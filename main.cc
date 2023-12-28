#include <curl/curl.h>
#include <gumbo.h>
#include <xlsxio_read.h>

#include <chrono>
#include <cstring>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>
#include <variant>
#include <vector>

struct Error {
    const std::string msg;

    [[noreturn]] void fatal() const {
        print();
        exit(1);
    }

    void print() const {
        const char redColor[6] = "\033[31m";
        const char resetColor[6] = "\033[0m";
        std::cerr << redColor << "Error: " << msg << resetColor << std::endl;
    }
};

template <typename T = std::monostate>
struct Result {
    const std::variant<Error, T> val;

    Result() : val(std::monostate()) {}
    Result(const T& val) : val(val) {}
    Result(const Error& error) : val(error) {}

    T result() const {
        if (success()) {
            return std::get<T>(val);
        } else {
            std::get<Error>(val).fatal();
        }
    }

    bool success() const {
        return std::holds_alternative<T>(val);
    }
};

struct ExcelConfig {
    const std::string fileName;
    const std::string sheetName;
    const int dateColIdx;
};

struct NewDataTime {
    const time_t firstMonthTime;
    const time_t presentMonthTime;
    const int firstMonthDay;
};

struct ResponseChunksBuffer {
    std::vector<char> buffer;

    void append(const char* chunk, size_t size) {
        buffer.insert(buffer.end(), chunk, chunk + size);
    }

    void clear() {
        buffer.clear();
    }

    const char* data() const {
        return buffer.data();
    }

    size_t size() const {
        return buffer.size();
    }
};

Result<ExcelConfig> getExcelConfig();
Result<int> stoiDateColIdx(char* val);

Result<NewDataTime> getNewDataTime(const ExcelConfig& excelConfig);
Result<NewDataTime> parseExcelDateStr(const std::string& dateStr);
time_t getPresentMonthTime();

Result<> requestNewWeatherData(NewDataTime newDataTime);
size_t curlWriteFunction(void* contents, size_t size, size_t nmemb, void* userp);
void parseHtmlAndWrite(const char* html);
void searchNodes(GumboNode* node);

Result<ExcelConfig> getExcelConfig() {
    xlsxioreader xlsxioReader;
    if ((xlsxioReader = xlsxioread_open("config.xlsx")) == NULL) {
        return Error{"Failed to open config.xlsx"};
    }

    std::string fileName;
    std::string sheetName;
    int dateColIdx;

    int col = 0;
    char* val;
    xlsxioreadersheet sheet = xlsxioread_sheet_open(xlsxioReader, "Config", XLSXIOREAD_SKIP_EMPTY_ROWS);
    xlsxioread_sheet_next_row(sheet);
    xlsxioread_sheet_next_row(sheet);
    while ((val = xlsxioread_sheet_next_cell(sheet)) != NULL) {
        switch (col) {
            case 0:
                fileName = val;
                break;
            case 1:
                sheetName = val;
                break;
            case 2:
                dateColIdx = stoiDateColIdx(val).result();
                break;
        }
        xlsxioread_free(val);
        col++;
    }

    if (fileName.empty()) {
        return Error{"Failed to  get EXCEL_FILE_NAME value"};
    }
    if (sheetName.empty()) {
        return Error{"Failed to get EXCEL_SHEET_NAME value"};
    }

    xlsxioread_sheet_close(sheet);
    xlsxioread_close(xlsxioReader);

    return ExcelConfig{fileName + ".xlsx", sheetName, dateColIdx};
}

Result<int> stoiDateColIdx(char* val) {
    try {
        return std::stoi(val);
    } catch (const std::invalid_argument& err) {
        return Error{"Invalid date column index"};
    }
}

Result<NewDataTime> getNewDataTime(const ExcelConfig& excelConfig) {
    xlsxioreader xlsxioReader;
    if ((xlsxioReader = xlsxioread_open(excelConfig.fileName.c_str())) == NULL) {
        return Error{"Failed to open " + excelConfig.fileName};
    }

    int col;
    char* val;
    char* lastDateVal = nullptr;
    xlsxioreadersheet sheet = xlsxioread_sheet_open(xlsxioReader, excelConfig.sheetName.c_str(), XLSXIOREAD_SKIP_EMPTY_ROWS);
    xlsxioread_sheet_next_row(sheet);
    while (xlsxioread_sheet_next_row(sheet)) {
        col = 0;
        while ((val = xlsxioread_sheet_next_cell(sheet)) != NULL) {
            if (col == excelConfig.dateColIdx) {
                free(lastDateVal);
                lastDateVal = strdup(val);
                break;
            }
            xlsxioread_free(val);
            col++;
        }
    }

    if (lastDateVal == nullptr) {
        return Error{"Failed to get last date value"};
    }

    xlsxioread_sheet_close(sheet);
    xlsxioread_close(xlsxioReader);
    return parseExcelDateStr(lastDateVal).result();
}

Result<NewDataTime> parseExcelDateStr(const std::string& dateStr) {
    std::tm date = {};
    std::istringstream ss(dateStr);
    ss >> std::get_time(&date, "%d.%m.%Y");
    if (ss.fail()) {
        return Error{"Failed to parse date: " + dateStr};
    }

    std::time_t firstMonthTime = std::mktime(&date) + (24 * 60 * 60);
    int firstMonthDay = date.tm_mday;
    int presentMonthTime = getPresentMonthTime();

    return NewDataTime{firstMonthTime, presentMonthTime, firstMonthDay};
}

time_t getPresentMonthTime() {
    auto chronoNow = std::chrono::system_clock::now();
    std::time_t nowTime = std::chrono::system_clock::to_time_t(chronoNow);
    std::tm now = *std::localtime(&nowTime);
    now.tm_sec = 0;
    now.tm_min = 0;
    now.tm_hour = 0;
    now.tm_wday = 0;
    now.tm_yday = 0;
    now.tm_isdst = 0;
    return std::mktime(&now);
}

Result<> requestNewWeatherData(NewDataTime newDataTime) {
    CURL* curl = curl_easy_init();
    if (!curl) {
        return Error{"Failed to initialize CURL"};
    }

    struct ResponseChunksBuffer responseChunksBuffer;
    curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, curlWriteFunction);
    curl_easy_setopt(curl, CURLOPT_WRITEDATA, &responseChunksBuffer);

    time_t monthTime = newDataTime.firstMonthTime;
    while (monthTime <= newDataTime.presentMonthTime) {
        struct tm* monthDate = localtime(&monthTime);
        std::ostringstream urlStream;
        urlStream << "https://freemeteo.ro/vremea/bucuroaia/istoric/istoric-lunar/?gid=683499&station=4621"
                  << "&month=" << monthDate->tm_mon + 1
                  << "&year=" << monthDate->tm_year + 1900
                  << "&language=romanian&country=romania";
        std::string url = urlStream.str();
        curl_easy_setopt(curl, CURLOPT_URL, url.c_str());
        CURLcode res = curl_easy_perform(curl);

        if (res != CURLE_OK) {
            return Error{"CURL request failed: " + std::string(curl_easy_strerror(res))};
        }

        parseHtmlAndWrite(responseChunksBuffer.data());
        responseChunksBuffer.clear();

        monthTime += (30 * 24 * 60 * 60);
    }

    curl_easy_cleanup(curl);

    return {};
}

size_t curlWriteFunction(void* contents, size_t size, size_t nmemb, void* userp) {
    size_t addSize = size * nmemb;
    ResponseChunksBuffer* buffer = static_cast<ResponseChunksBuffer*>(userp);
    buffer->append(static_cast<const char*>(contents), addSize);
    return addSize;
}

void parseHtmlAndWrite(const char* html) {
    GumboOutput* output = gumbo_parse(html);
    searchNodes(output->root);
    gumbo_destroy_output(&kGumboDefaultOptions, output);
}

void searchNodes(GumboNode* node) {
    if (node->type != GUMBO_NODE_ELEMENT) {
        return;
    }

    GumboAttribute* attr;
    if (node->v.element.tag == GUMBO_TAG_TR && (attr = gumbo_get_attribute(&node->v.element.attributes, "data-day"))) {
        GumboVector* tdNodes = &node->v.element.children;

        // start with 1 and increment by 2 to jump over nodes of type GUMBO_NODE_WHITESPACE
        for (unsigned int i = 1; i < tdNodes->length; i += 2) {
            GumboNode* tdNode = static_cast<GumboNode*>(tdNodes->data[i]);
            GumboNode* textNode;

            if (i == 1) {
                GumboNode* anchorNode = static_cast<GumboNode*>(tdNode->v.element.children.data[0]);
                textNode = static_cast<GumboNode*>(anchorNode->v.element.children.data[0]);
                std::cout << textNode->v.text.text << std::endl;
            } else if (i != 17) {
                textNode = static_cast<GumboNode*>(tdNode->v.element.children.data[0]);
                std::cout << textNode->v.text.text << std::endl;
            }
        }

        std::cout << std::endl;
    }

    GumboVector* children = &node->v.element.children;
    for (unsigned int i = 0; i < children->length; ++i) {
        searchNodes(static_cast<GumboNode*>(children->data[i]));
    }
}

int main() {
    ExcelConfig excelConfig = getExcelConfig().result();

    std::cout << excelConfig.fileName << std::endl;
    std::cout << excelConfig.sheetName << std::endl;
    std::cout << excelConfig.dateColIdx << std::endl
              << std::endl;

    NewDataTime newDataTime = getNewDataTime(excelConfig).result();

    std::cout << newDataTime.firstMonthTime << std::endl;
    std::cout << newDataTime.presentMonthTime << std::endl;
    std::cout << newDataTime.firstMonthDay << std::endl
              << std::endl;

    requestNewWeatherData(newDataTime).result();

    return 0;
}
