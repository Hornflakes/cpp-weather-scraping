#include <curl/curl.h>
#include <gumbo.h>
#include <stdio.h>
#include <string.h>
#include <time.h>
#include <xlsxio_read.h>

struct DataPointKeyIndex {
    int date;
    int minTemperature;
    int maxTemperature;
    int maxSustainedWind;
    int maxGustWind;
    int rainfall;
    int snowDepth;
    int description;
};

struct ResponseChunksBuffer {
    char* memory;
    size_t size;
};

const char* EXCEL_FILE_NAME;
const char* EXCEL_SHEET_NAME;
int DATE_COL_EXCEL_INDEX;
struct DataPointKeyIndex DATA_POINT_KEY_INDEX;
time_t FIRST_MONTH_TIME;
time_t PRESENT_MONTH_TIME;
int FIRST_DAY_DATE;

void readConfigExcel();
int dateColExcelIndex(char* value);
struct DataPointKeyIndex dataPointKeyIndex(int index);

void readWeatherExcel();
struct tm parseExcelDate(const char* dateString);
void addDayToTm(struct tm* date);
time_t presentMonthTime();

void requestWeatherData();
size_t curlWriteFunction(void* contents, size_t size, size_t nmemb, void* userp);
int adjustMonth(int month);
int adjustYear(int year);
void parseAndWrite(const char* html);
void searchNodes(GumboNode* node);

void readConfigExcel() {
    xlsxioreader xlsxioReader;
    if ((xlsxioReader = xlsxioread_open("config.xlsx")) == NULL) {
        fprintf(stderr, "Error opening config.xlsx file\n");
        exit(1);
    }

    int col = 0;
    char* value;
    xlsxioreadersheet sheet = xlsxioread_sheet_open(xlsxioReader, "Config", XLSXIOREAD_SKIP_EMPTY_ROWS);
    xlsxioread_sheet_next_row(sheet);
    xlsxioread_sheet_next_row(sheet);
    while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
        switch (col) {
            case 0:
                EXCEL_FILE_NAME = strcat(strdup(value), ".xlsx");
                break;
            case 1:
                EXCEL_SHEET_NAME = strdup(value);
                break;
            case 2:
                DATE_COL_EXCEL_INDEX = dateColExcelIndex(value);
                DATA_POINT_KEY_INDEX = dataPointKeyIndex(DATE_COL_EXCEL_INDEX);
                break;
        }
        xlsxioread_free(value);
        col++;
    }

    if (EXCEL_FILE_NAME == NULL) {
        fprintf(stderr, "Error opening Config sheet\n");
        exit(1);
    }
    if (strlen(EXCEL_FILE_NAME) == 0) {
        fprintf(stderr, "Error getting EXCEL_FILE_NAME value\n");
        exit(1);
    }
    if (strlen(EXCEL_SHEET_NAME) == 0) {
        fprintf(stderr, "Error getting EXCEL_SHEET_NAME value\n");
        exit(1);
    }

    printf("EXCEL_FILE_NAME: %s\n", EXCEL_FILE_NAME);
    printf("EXCEL_SHEET_NAME: %s\n", EXCEL_SHEET_NAME);
    printf("DATE_COL_EXCEL_INDEX: %i\n", DATE_COL_EXCEL_INDEX);

    xlsxioread_sheet_close(sheet);
    xlsxioread_close(xlsxioReader);
}

int dateColExcelIndex(char* value) {
    char* endPtr;
    int dateColExcelIndex = strtol(value, &endPtr, 10);

    if (endPtr == value || *endPtr != '\0') {
        fprintf(stderr, "Error getting DATE_COL_EXCEL_INDEX value\n");
        exit(1);
    }

    return dateColExcelIndex;
}

struct DataPointKeyIndex dataPointKeyIndex(int index) {
    struct DataPointKeyIndex dataPointKeyIndex;
    dataPointKeyIndex.date = index++;
    dataPointKeyIndex.minTemperature = index++;
    dataPointKeyIndex.maxTemperature = index++;
    dataPointKeyIndex.maxSustainedWind = index++;
    dataPointKeyIndex.maxGustWind = index++;
    dataPointKeyIndex.rainfall = index++;
    dataPointKeyIndex.snowDepth = index++;
    dataPointKeyIndex.description = index++;
    return dataPointKeyIndex;
}

void readWeatherExcel() {
    xlsxioreader xlsxioReader;
    if ((xlsxioReader = xlsxioread_open(EXCEL_FILE_NAME)) == NULL) {
        fprintf(stderr, "Error opening %s file\n", EXCEL_FILE_NAME);
        exit(1);
    }

    int col;
    char* value;
    char* lastValue;
    xlsxioreadersheet sheet = xlsxioread_sheet_open(xlsxioReader, EXCEL_SHEET_NAME, XLSXIOREAD_SKIP_EMPTY_ROWS);
    xlsxioread_sheet_next_row(sheet);
    while (xlsxioread_sheet_next_row(sheet)) {
        col = 0;
        while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
            if (col == DATE_COL_EXCEL_INDEX) {
                free(lastValue);
                lastValue = strdup(value);
                break;
            }
            xlsxioread_free(value);
            col++;
        }
    }

    struct tm firstDayDate = parseExcelDate(lastValue);
    addDayToTm(&firstDayDate);

    FIRST_MONTH_TIME = mktime(&firstDayDate);
    PRESENT_MONTH_TIME = presentMonthTime();
    FIRST_DAY_DATE = firstDayDate.tm_mday;

    xlsxioread_sheet_close(sheet);
    xlsxioread_close(xlsxioReader);
}

struct tm parseExcelDate(const char* dateString) {
    if (strlen(dateString) != 10 || dateString[2] != '.' || dateString[5] != '.') {
        printf("Error parsing last date value\n");
        exit(1);
    }

    struct tm date;
    date.tm_mday = atoi(dateString);
    date.tm_mon = atoi(dateString + 3) - 1;
    date.tm_year = atoi(dateString + 6) - 1900;
    date.tm_sec = 0;
    date.tm_min = 0;
    date.tm_hour = 0;
    date.tm_wday = 0;
    date.tm_yday = 0;
    date.tm_isdst = 0;
    return date;
}

void addDayToTm(struct tm* date) {
    time_t t = mktime(date);
    t += 86400;
    *date = *localtime(&t);
}

time_t presentMonthTime() {
    time_t rawTime = time(NULL);
    struct tm* presentMonthDate = localtime(&rawTime);
    presentMonthDate->tm_sec = 0;
    presentMonthDate->tm_min = 0;
    presentMonthDate->tm_hour = 0;
    presentMonthDate->tm_wday = 0;
    presentMonthDate->tm_yday = 0;
    presentMonthDate->tm_isdst = 0;
    return mktime(presentMonthDate);
}

void requestWeatherData() {
    CURL* curl = curl_easy_init();
    if (!curl) {
        fprintf(stderr, "Error initializing curl\n");
        exit(1);
    }

    struct ResponseChunksBuffer responseChunksBuffer = {
        .memory = NULL,
        .size = 0};
    curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, curlWriteFunction);
    curl_easy_setopt(curl, CURLOPT_WRITEDATA, (void*)&responseChunksBuffer);

    time_t monthTime = FIRST_MONTH_TIME;
    while (monthTime <= PRESENT_MONTH_TIME) {
        struct tm* monthDate = localtime(&monthTime);
        char url[256];
        snprintf(url, sizeof(url), "https://freemeteo.ro/vremea/bucuroaia/istoric/istoric-lunar/?gid=683499&station=4621&month=%d&year=%d&language=romanian&country=romania", adjustMonth(monthDate->tm_mon), adjustYear(monthDate->tm_year));

        curl_easy_setopt(curl, CURLOPT_URL, url);
        CURLcode res = curl_easy_perform(curl);

        if (res != CURLE_OK) {
            fprintf(stderr, "CURL request failed: %s\n", curl_easy_strerror(res));
            free(responseChunksBuffer.memory);
            exit(1);
        }

        parseAndWrite(responseChunksBuffer.memory);

        free(responseChunksBuffer.memory);
        responseChunksBuffer.size = 0;

        monthTime += 86400 * 300;
    }

    curl_easy_cleanup(curl);
}

size_t curlWriteFunction(void* contents, size_t size, size_t nmemb, void* userp) {
    size_t sizeAdd = size * nmemb;

    struct ResponseChunksBuffer* buffer = (struct ResponseChunksBuffer*)userp;
    char* ptr = realloc(buffer->memory, buffer->size + sizeAdd + 1);
    buffer->memory = ptr;
    memcpy(&(buffer->memory[buffer->size]), contents, sizeAdd);
    buffer->size += sizeAdd;
    buffer->memory[buffer->size] = 0;

    return sizeAdd;
}

int adjustMonth(int month) {
    return month + 1;
}

int adjustYear(int year) {
    return year + 1900;
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
            GumboNode* tdNode = tdNodes->data[i];
            GumboNode* textNode;

            if (i == 1) {
                GumboNode* anchorNode = tdNode->v.element.children.data[0];
                textNode = anchorNode->v.element.children.data[0];
                printf("%s\n", textNode->v.text.text);
            } else if (i != 17) {
                textNode = tdNode->v.element.children.data[0];
                printf("%s\n", textNode->v.text.text);
            }
        }

        printf("\n");
    }

    GumboVector* children = &node->v.element.children;
    for (unsigned int i = 0; i < children->length; ++i) {
        searchNodes(children->data[i]);
    }
}

void parseAndWrite(const char* html) {
    GumboOutput* output = gumbo_parse(html);
    searchNodes(output->root);
    gumbo_destroy_output(&kGumboDefaultOptions, output);
    free(output);
}

int main() {
    readConfigExcel();
    readWeatherExcel();
    requestWeatherData();

    free((void*)EXCEL_FILE_NAME);
    free((void*)EXCEL_SHEET_NAME);

    return 0;
}
