#include <curl/curl.h>
#include <gumbo.h>

#include <chrono>
#include <cstring>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>
#include <variant>
#include <vector>
#include <xlnt/xlnt.hpp>

struct Error {
    const std::string msg;

    [[noreturn]] void fatal() const {
        print();
        exit(1);
    }

    void print() const {
        const char redColor[6] = "\033[31m";
        const char resetColor[6] = "\033[0m";
        std::cerr << redColor << "Error : " << msg << resetColor << std::endl;
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

struct NewDataParams {
    const NewDataTime newDataTime;
    const int startRowIdx;
};

struct WeatherDataPoint {
    const std::string date;
    const std::string minTemperature;
    const std::string maxTemperature;
    const std::string maxSustainedWind;
    const std::string maxGustWind;
    const std::string rainfall;
    const std::string snowdepth;
    const std::string description;
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
Result<int> stoiDateColIdx(std::string val);

Result<NewDataParams> getNewDataParams(const ExcelConfig& excelConfig);
Result<NewDataTime> parseExcelDateStr(const std::string& dateStr);
time_t getPresentMonthTime();

Result<std::vector<WeatherDataPoint>> getWeatherData(NewDataParams newDataParams);
size_t curlWriteFunction(void* contents, size_t size, size_t nmemb, void* userp);
Result<std::vector<WeatherDataPoint>> getMonthlyWeatherData(const char* html);
void parseHtml(GumboNode* node, std::vector<WeatherDataPoint>& monthlyWeatherData);
Result<WeatherDataPoint> getWeatherDataPoint(GumboNode* node);

Result<ExcelConfig> getExcelConfig() {
    xlnt::workbook wb;
    try {
        wb.load("config.xlsx");
    } catch (const xlnt::exception& err) {
        return Error{"Failed to open config.xlsx : " + std::string(err.what())};
    }

    xlnt::worksheet ws;
    try {
        ws = wb.sheet_by_title("Config");
    } catch (const xlnt::exception& err) {
        return Error{"Failed to get sheet Config : " + std::string(err.what())};
    }

    std::string fileName;
    std::string sheetName;
    int dateColIdx;
    try {
        fileName = ws.cell(xlnt::cell_reference("A2")).to_string();
    } catch (const xlnt::exception& err) {
        return Error{"Failed to get EXCEL_FILE_NAME value : " + std::string(err.what())};
    }
    try {
        sheetName = ws.cell(xlnt::cell_reference("B2")).to_string();
    } catch (const xlnt::exception& err) {
        return Error{"Failed to get EXCEL_SHEET_NAME value : " + std::string(err.what())};
    }
    try {
        dateColIdx = stoiDateColIdx(ws.cell(xlnt::cell_reference("C2")).to_string()).result();
    } catch (const xlnt::exception& err) {
        return Error{"Failed to get EXCEL_DATE_COL_IDX value : " + std::string(err.what())};
    }

    if (fileName.empty()) {
        return Error{"EXCEL_FILE_NAME value cannot be empty"};
    }
    if (sheetName.empty()) {
        return Error{"EXCEL_SHEET_NAME value cannot be empty"};
    }

    return ExcelConfig{fileName + ".xlsx", sheetName, dateColIdx};
}

Result<int> stoiDateColIdx(std::string val) {
    try {
        return std::stoi(val);
    } catch (const std::invalid_argument& err) {
        return Error{"Invalid date column index"};
    }
}

Result<NewDataParams> getNewDataParams(const ExcelConfig& excelConfig) {
    xlnt::workbook wb;
    try {
        wb.load(excelConfig.fileName);
    } catch (const xlnt::exception& err) {
        return Error{"Failed to open " + excelConfig.fileName + " : " + std::string(err.what())};
    }

    xlnt::worksheet ws;
    try {
        ws = wb.sheet_by_title(excelConfig.sheetName);
    } catch (const xlnt::exception& err) {
        return Error{"Failed to open sheet " + excelConfig.sheetName + " : " + std::string(err.what())};
    }

    std::string lastDateStr;
    int startingRowIdx = 1;
    for (auto row : ws.rows()) {
        if (!row[excelConfig.dateColIdx].to_string().empty()) {
            lastDateStr = row[excelConfig.dateColIdx].to_string();
        }
        ++startingRowIdx;
    }

    if (lastDateStr.empty()) {
        return Error{"Last date value not found"};
    }

    NewDataTime newDataTime = parseExcelDateStr(lastDateStr).result();
    return NewDataParams{newDataTime, startingRowIdx};
}

Result<NewDataTime> parseExcelDateStr(const std::string& dateStr) {
    std::tm date = {};
    std::istringstream ss(dateStr);
    ss >> std::get_time(&date, "%d.%m.%Y");
    if (ss.fail()) {
        return Error{"Failed to parse date : " + dateStr};
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

Result<std::vector<WeatherDataPoint>> getWeatherData(NewDataParams newDataParams) {
    std::vector<WeatherDataPoint> weatherData;

    CURL* curl = curl_easy_init();
    if (!curl) {
        return Error{"Failed to initialize CURL"};
    }

    struct ResponseChunksBuffer responseChunksBuffer;
    curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, curlWriteFunction);
    curl_easy_setopt(curl, CURLOPT_WRITEDATA, &responseChunksBuffer);

    time_t monthTime = newDataParams.newDataTime.firstMonthTime;
    while (monthTime <= newDataParams.newDataTime.presentMonthTime) {
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
            return Error{"CURL request failed : " + std::string(curl_easy_strerror(res))};
        }

        std::vector<WeatherDataPoint> monthlyWeatherData = getMonthlyWeatherData(responseChunksBuffer.data()).result();
        for (WeatherDataPoint dataPoint : monthlyWeatherData) {
            weatherData.push_back(dataPoint);
        }
        responseChunksBuffer.clear();

        monthTime += (30 * 24 * 60 * 60);
    }

    curl_easy_cleanup(curl);

    return weatherData;
}

size_t curlWriteFunction(void* contents, size_t size, size_t nmemb, void* userp) {
    size_t addSize = size * nmemb;
    ResponseChunksBuffer* buffer = static_cast<ResponseChunksBuffer*>(userp);
    buffer->append(static_cast<const char*>(contents), addSize);
    return addSize;
}

Result<std::vector<WeatherDataPoint>> getMonthlyWeatherData(const char* html) {
    GumboOutput* output = gumbo_parse(html);
    if (!output) {
        return Error{"Failed to parse HTML"};
    }

    std::vector<WeatherDataPoint> monthlyWeatherData;
    parseHtml(output->root, monthlyWeatherData);

    gumbo_destroy_output(&kGumboDefaultOptions, output);

    return monthlyWeatherData;
}

void parseHtml(GumboNode* node, std::vector<WeatherDataPoint>& monthlyWeatherData) {
    if (!node || node->type != GUMBO_NODE_ELEMENT) {
        return;
    }

    if (node->v.element.tag == GUMBO_TAG_TR) {
        GumboAttribute* attr = gumbo_get_attribute(&node->v.element.attributes, "data-day");
        if (!attr) {
            return;
        }

        WeatherDataPoint weatherDataPoint = getWeatherDataPoint(node).result();
        monthlyWeatherData.push_back(weatherDataPoint);
    }

    GumboVector* children = &node->v.element.children;
    for (unsigned int i = 0; i < children->length; ++i) {
        parseHtml(static_cast<GumboNode*>(children->data[i]), monthlyWeatherData);
    }
}

Result<WeatherDataPoint> getWeatherDataPoint(GumboNode* node) {
    std::string date;
    std::string minTemperature;
    std::string maxTemperature;
    std::string maxSustainedWind;
    std::string maxGustWind;
    std::string rainfall;
    std::string snowdepth;
    std::string description;

    GumboVector* tdNodes = &node->v.element.children;
    // start with 1 and increment by 2 to jump over nodes of type GUMBO_NODE_WHITESPACE
    for (unsigned int i = 1; i < tdNodes->length; i += 2) {
        if (i == 15 || i == 17) {
            continue;
        }

        GumboNode* tdNode = static_cast<GumboNode*>(tdNodes->data[i]);
        if (!tdNode) {
            return Error{"Failed to get TD node, website structure might have changed"};
        }

        GumboNode* textNode;
        // the date is an anchor node
        if (i == 1) {
            GumboNode* anchorNode = static_cast<GumboNode*>(tdNode->v.element.children.data[0]);
            textNode = static_cast<GumboNode*>(anchorNode->v.element.children.data[0]);
        } else {
            textNode = static_cast<GumboNode*>(tdNode->v.element.children.data[0]);
        }

        if (!textNode) {
            return Error{"Failed to get text, website structure might have changed"};
        }

        std::string text = textNode->v.text.text;
        switch (i) {
            case 1:
                date = text;
                break;
            case 3:
                minTemperature = text;
                break;
            case 5:
                maxTemperature = text;
                break;
            case 7:
                maxSustainedWind = text;
                break;
            case 9:
                maxGustWind = text;
                break;
            case 11:
                rainfall = text;
                break;
            case 13:
                snowdepth = text;
                break;
            case 19:
                description = text;
                break;
        }
    }

    return WeatherDataPoint{date, minTemperature, maxTemperature, maxSustainedWind, maxGustWind, rainfall, snowdepth, description};
}

int main() {
    ExcelConfig excelConfig = getExcelConfig().result();

    std::cout << excelConfig.fileName << std::endl;
    std::cout << excelConfig.sheetName << std::endl;
    std::cout << excelConfig.dateColIdx << std::endl
              << std::endl;

    NewDataParams newDataParams = getNewDataParams(excelConfig).result();

    std::cout << newDataParams.newDataTime.firstMonthTime << std::endl;
    std::cout << newDataParams.newDataTime.presentMonthTime << std::endl;
    std::cout << newDataParams.newDataTime.firstMonthDay << std::endl
              << std::endl;

    std::vector<WeatherDataPoint> weatherData = getWeatherData(newDataParams).result();

    for (WeatherDataPoint data : weatherData) {
        std::cout << data.date << std::endl;
        std::cout << data.minTemperature << std::endl;
        std::cout << data.maxTemperature << std::endl;
        std::cout << data.maxSustainedWind << std::endl;
        std::cout << data.maxGustWind << std::endl;
        std::cout << data.rainfall << std::endl;
        std::cout << data.snowdepth << std::endl;
        std::cout << data.description << std::endl
                  << std::endl;
    }

    return 0;
}
