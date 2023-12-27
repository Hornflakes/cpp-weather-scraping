#include <stdio.h>
#include <string.h>
#include <xlsxio_read.h>

const char* EXCEL_FILE_NAME;
const char* EXCEL_SHEET_NAME;
const char* DATE_COL_EXCEL_INDEX;

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
                EXCEL_FILE_NAME = strdup(value);
                break;
            case 1:
                EXCEL_SHEET_NAME = strdup(value);
                break;
            case 2:
                DATE_COL_EXCEL_INDEX = strdup(value);
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
    if (strlen(DATE_COL_EXCEL_INDEX) == 0) {
        fprintf(stderr, "Error getting DATE_COL_EXCEL_INDEX value\n");
        exit(1);
    }

    printf("EXCEL_FILE_NAME: %s\n", EXCEL_FILE_NAME);
    printf("EXCEL_SHEET_NAME: %s\n", EXCEL_SHEET_NAME);
    printf("DATE_COL_EXCEL_INDEX: %s\n", DATE_COL_EXCEL_INDEX);

    xlsxioread_sheet_close(sheet);
    xlsxioread_close(xlsxioReader);
}

int main() {
    readConfigExcel();

    free((void*)EXCEL_FILE_NAME);
    free((void*)EXCEL_SHEET_NAME);
    free((void*)DATE_COL_EXCEL_INDEX);

    return 0;
}
