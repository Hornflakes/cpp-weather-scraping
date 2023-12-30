#include <curl/curl.h>
#include <gumbo.h>

#include <algorithm>
#include <chrono>
#include <cstring>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>
#include <variant>
#include <vector>
#include <xlnt/xlnt.hpp>

#ifdef _WIN32
#include <windows.h>
#endif

struct Error {
    const std::string msg;

    [[noreturn]] void fatal() const {
        print();
        exit(1);
    }

#ifdef _WIN32
    void print() const {
        setColor(FOREGROUND_RED);
        std::cerr << "Error : " << msg << std::endl;
        setColor();
    }

    void setColor(int color = FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE) const {
        HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
        SetConsoleTextAttribute(hConsole, color);
    }
#else
    void print() const {
        const char redColor[6] = "\033[31m";
        const char resetColor[6] = "\033[0m";
        std::cerr << redColor << "Error : " << msg << resetColor << std::endl;
    }
#endif
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
    const std::string dateColumnLetter;
};

struct NewDataTime {
    const time_t firstMonthTime;
    const time_t presentMonthTime;
    const unsigned short firstMonthDay;
};

struct NewDataParams {
    const NewDataTime newDataTime;
    const unsigned short startRowIdx;
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

Result<NewDataParams> getNewDataParams(ExcelConfig excelConfig);
Result<NewDataTime> parseExcelDateStr(const std::string dateStr);
std::tm normalizedDateTime(std::tm dateTime);
time_t getPresentMonthTime();

Result<std::vector<WeatherDataPoint>> getWeatherData(NewDataParams newDataParams);
time_t addOneMonth(time_t time);
size_t curlWriteFunction(void* contents, size_t size, size_t nmemb, void* userp);
Result<std::vector<WeatherDataPoint>> getMonthlyWeatherData(const char* html, bool isFirstMonth, NewDataTime newDataTime);
void parseHtml(GumboNode* node, std::vector<WeatherDataPoint>& monthlyWeatherData, bool isFirstMonth, NewDataTime newDataTime);
Result<WeatherDataPoint> getWeatherDataPoint(GumboNode* node);
std::string quoteAfterNegativeNumber(std::string& str);

Result<> writeWeatherExcel(ExcelConfig excelConfig, NewDataParams newDataParams, std::vector<WeatherDataPoint>& weatherData);

Result<ExcelConfig> getExcelConfig() {
    xlnt::workbook wb;
    try {
        wb.load("config.xlsx");
    } catch (const std::exception& err) {
        return Error{"Failed to open config.xlsx : " + std::string(err.what()) + "\nMake sure file config.xlsx exists or is in the same folder as the executable"};
    }

    xlnt::worksheet ws;
    try {
        ws = wb.sheet_by_title("Config");
    } catch (const std::exception& err) {
        return Error{"Failed to get sheet Config : " + std::string(err.what()) + "\nMake sure the sheet is named Config"};
    }

    std::string fileName;
    std::string sheetName;
    std::string dateColumnLetter;
    try {
        fileName = ws.cell(xlnt::cell_reference("A2")).to_string();
    } catch (const std::exception& err) {
        return Error{"Failed to get EXCEL_FILE_NAME value : " + std::string(err.what())};
    }
    try {
        sheetName = ws.cell(xlnt::cell_reference("B2")).to_string();
    } catch (const xlnt::exception& err) {
        return Error{"Failed to get EXCEL_SHEET_NAME value : " + std::string(err.what())};
    }
    try {
        dateColumnLetter = ws.cell(xlnt::cell_reference("C2")).to_string();
    } catch (const xlnt::exception& err) {
        return Error{"Failed to get DATE_COLUMN_LETTER value : " + std::string(err.what())};
    }

    if (fileName.empty()) {
        return Error{"EXCEL_FILE_NAME value cannot be empty\nMake sure the value is in cell A2 of the Config sheet"};
    }
    if (sheetName.empty()) {
        return Error{"EXCEL_SHEET_NAME value cannot be empty\nMake sure the value is in cell B2 of the Config sheet"};
    }
    if (dateColumnLetter.empty()) {
        return Error{"DATE_COLUMN_LETTER value cannot be empty\nMake sure the value is in cell C2 of the Config sheet"};
    }
    if (std::any_of(dateColumnLetter.begin(), dateColumnLetter.end(), ::isdigit)) {
        return Error{"DATE_COLUMN_LETTER cannot contain numbers"};
    }

    return ExcelConfig{fileName + ".xlsx", sheetName, dateColumnLetter};
}

Result<NewDataParams> getNewDataParams(ExcelConfig excelConfig) {
    xlnt::workbook wb;
    try {
        wb.load(excelConfig.fileName);
    } catch (const std::exception& err) {
        return Error{"Failed to open " + excelConfig.fileName + " : " + std::string(err.what()) + "\nMake sure file " + excelConfig.fileName + " exists or is in the same folder as the executable"};
    }

    xlnt::worksheet ws;
    try {
        ws = wb.sheet_by_title(excelConfig.sheetName);
    } catch (const std::exception& err) {
        return Error{"Failed to open sheet " + excelConfig.sheetName + " : " + std::string(err.what()) + "\nMake sure sheet " + excelConfig.sheetName + " exists"};
    }

    std::string lastDateStr;
    unsigned short rowIdx = ws.highest_row();
    for (; rowIdx > 1; --rowIdx) {
        xlnt::cell cell = ws.cell(excelConfig.dateColumnLetter + std::to_string(rowIdx));
        std::string val = cell.to_string();
        if (!val.empty()) {
            lastDateStr = val;
            break;
        }
    }

    if (lastDateStr.empty()) {
        return Error{"Last date value not found\nMake sure column " + excelConfig.dateColumnLetter + " has a value"};
    }

    NewDataTime newDataTime = parseExcelDateStr(lastDateStr).result();
    return NewDataParams{newDataTime, ++rowIdx};
}

Result<NewDataTime> parseExcelDateStr(const std::string dateStr) {
    std::tm dateTime = {};
    std::istringstream ss(dateStr);
    ss >> std::get_time(&dateTime, "%d.%m.%Y");
    if (ss.fail()) {
        return Error{"Failed to parse date : " + dateStr + "\nMake sure the date is in format DD.MM.YYYY"};
    }
    dateTime = normalizedDateTime(dateTime);

    std::time_t firstMonthTime = std::mktime(&dateTime) + (24 * 60 * 60);
    std::tm firstMonthDateTime = *std::localtime(&firstMonthTime);
    unsigned short firstMonthDay = firstMonthDateTime.tm_mday;
    // firstMonthDateTime set to first day of the month to later add months without edge cases
    firstMonthDateTime.tm_mday = 1;
    time_t presentMonthTime = getPresentMonthTime();

    return NewDataTime{firstMonthTime, presentMonthTime, firstMonthDay};
}

std::tm normalizedDateTime(std::tm dateTime) {
    dateTime.tm_sec = 0;
    dateTime.tm_min = 0;
    dateTime.tm_hour = 0;
    dateTime.tm_wday = 0;
    dateTime.tm_yday = 0;
    dateTime.tm_isdst = 0;
    return dateTime;
}

time_t getPresentMonthTime() {
    auto chronoNow = std::chrono::system_clock::now();
    std::time_t nowTime = std::chrono::system_clock::to_time_t(chronoNow);
    std::tm nowDateTime = normalizedDateTime(*std::localtime(&nowTime));
    return std::mktime(&nowDateTime);
}

Result<std::vector<WeatherDataPoint>> getWeatherData(NewDataParams newDataParams) {
    std::vector<WeatherDataPoint> weatherData;

    CURL* curl = curl_easy_init();
    if (!curl) {
        return Error{"Failed to initialize CURL"};
    }

    ResponseChunksBuffer responseChunksBuffer;
    curl_easy_setopt(curl, CURLOPT_CAINFO, "./certs/cacert-2023-12-12.pem");
    curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, curlWriteFunction);
    curl_easy_setopt(curl, CURLOPT_WRITEDATA, &responseChunksBuffer);

    std::ostringstream urlStream;
    time_t monthTime = newDataParams.newDataTime.firstMonthTime;
    while (monthTime <= newDataParams.newDataTime.presentMonthTime) {
        std::tm monthDate = *localtime(&monthTime);
        urlStream << "https://freemeteo.ro/vremea/bucuroaia/istoric/istoric-lunar/?gid=683499&station=4621"
                  << "&month=" << monthDate.tm_mon + 1
                  << "&year=" << monthDate.tm_year + 1900
                  << "&language=romanian&country=romania";
        std::string url = urlStream.str();
        urlStream.str("");
        urlStream.clear();

        curl_easy_setopt(curl, CURLOPT_URL, url.c_str());
        CURLcode res = curl_easy_perform(curl);
        if (res != CURLE_OK) {
            return Error{"CURL request failed : " + std::string(curl_easy_strerror(res))};
        }

        bool isFirstMonth = monthTime == newDataParams.newDataTime.firstMonthTime;
        std::vector<WeatherDataPoint> monthlyWeatherData = getMonthlyWeatherData(responseChunksBuffer.data(), isFirstMonth, newDataParams.newDataTime).result();
        for (WeatherDataPoint dataPoint : monthlyWeatherData) {
            weatherData.push_back(dataPoint);
        }
        responseChunksBuffer.clear();

        monthTime = addOneMonth(monthTime);
    }

    curl_easy_cleanup(curl);

    return weatherData;
}

time_t addOneMonth(time_t time) {
    std::tm dateTime = *std::localtime(&time);
    dateTime.tm_mon += 1;
    if (dateTime.tm_mon > 11) {
        dateTime.tm_mon = 0;
        dateTime.tm_year += 1;
    }
    return std::mktime(&dateTime);
}

size_t curlWriteFunction(void* contents, size_t size, size_t nmemb, void* userp) {
    size_t addSize = size * nmemb;
    ResponseChunksBuffer* buffer = static_cast<ResponseChunksBuffer*>(userp);
    buffer->append(static_cast<const char*>(contents), addSize);
    return addSize;
}

Result<std::vector<WeatherDataPoint>> getMonthlyWeatherData(const char* html, bool isFirstMonth, NewDataTime newDataTime) {
    GumboOutput* output = gumbo_parse(html);
    if (!output) {
        return Error{"Failed to parse HTML"};
    }

    std::vector<WeatherDataPoint> monthlyWeatherData;
    monthlyWeatherData.reserve(31);
    parseHtml(output->root, monthlyWeatherData, isFirstMonth, newDataTime);

    gumbo_destroy_output(&kGumboDefaultOptions, output);

    return monthlyWeatherData;
}

void parseHtml(GumboNode* node, std::vector<WeatherDataPoint>& monthlyWeatherData, bool isFirstMonth, NewDataTime newDataTime) {
    if (!node || node->type != GUMBO_NODE_ELEMENT) {
        return;
    }

    if (node->v.element.tag == GUMBO_TAG_TR) {
        GumboAttribute* attr = gumbo_get_attribute(&node->v.element.attributes, "data-day");
        if (!attr) {
            return;
        }
        if (isFirstMonth && std::stoi(attr->value) < newDataTime.firstMonthDay) {
            return;
        }

        WeatherDataPoint weatherDataPoint = getWeatherDataPoint(node).result();
        monthlyWeatherData.push_back(weatherDataPoint);
    }

    GumboVector* children = &node->v.element.children;
    for (unsigned short i = 0; i < children->length; ++i) {
        parseHtml(static_cast<GumboNode*>(children->data[i]), monthlyWeatherData, isFirstMonth, newDataTime);
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
    for (unsigned short i = 1; i < tdNodes->length; i += 2) {
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
                minTemperature = quoteAfterNegativeNumber(text);
                break;
            case 5:
                maxTemperature = quoteAfterNegativeNumber(text);
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

std::string quoteAfterNegativeNumber(std::string& str) {
    if (str[0] == '-') str.append("'");
    return str;
}

Result<> writeWeatherExcel(ExcelConfig excelConfig, NewDataParams newDataParams, std::vector<WeatherDataPoint>& weatherData) {
    xlnt::workbook wb;
    try {
        wb.load(excelConfig.fileName);
    } catch (const std::exception& err) {
        return Error{"Failed to open " + excelConfig.fileName + " : " + std::string(err.what()) + "\nMake sure file " + excelConfig.fileName + " exists or is in the same folder as the executable"};
    }

    xlnt::worksheet ws;
    try {
        ws = wb.sheet_by_title(excelConfig.sheetName);
    } catch (const std::exception& err) {
        return Error{"Failed to open sheet " + excelConfig.sheetName + " : " + std::string(err.what()) + "\nMake sure sheet " + excelConfig.sheetName + " exists"};
    }

    unsigned short rowIdx = newDataParams.startRowIdx;
    try {
        for (WeatherDataPoint data : weatherData) {
            xlnt::column_t column(excelConfig.dateColumnLetter);
            ws.cell(xlnt::cell_reference(column++, rowIdx)).value(data.date);
            ws.cell(xlnt::cell_reference(column++, rowIdx)).value(data.minTemperature);
            ws.cell(xlnt::cell_reference(column++, rowIdx)).value(data.maxTemperature);
            ws.cell(xlnt::cell_reference(column++, rowIdx)).value(data.maxSustainedWind);
            ws.cell(xlnt::cell_reference(column++, rowIdx)).value(data.maxGustWind);
            ws.cell(xlnt::cell_reference(column++, rowIdx)).value(data.rainfall);
            ws.cell(xlnt::cell_reference(column++, rowIdx)).value(data.snowdepth);
            ws.cell(xlnt::cell_reference(column, rowIdx)).value(data.description);
            ++rowIdx;
        }
    } catch (const std::exception& err) {
        return Error{"Failed to write data at row " + std::to_string(rowIdx) + " : " + std::string(err.what())};
    }

    try {
        wb.save(excelConfig.fileName);
    } catch (const std::exception& err) {
        return Error{"Failed to save " + excelConfig.fileName + " : " + std::string(err.what()) + "\nMake sure  the file " + excelConfig.fileName + " is not open"};
    }

    return {};
}

int main() {
    std::cout << "Weather scraping . . ." << std::endl;
    ExcelConfig excelConfig = getExcelConfig().result();
    NewDataParams newDataParams = getNewDataParams(excelConfig).result();
    std::vector<WeatherDataPoint> weatherData = getWeatherData(newDataParams).result();
    writeWeatherExcel(excelConfig, newDataParams, weatherData).result();
    return 0;
}
