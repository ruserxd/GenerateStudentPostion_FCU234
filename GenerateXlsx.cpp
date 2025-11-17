#include <iostream>
#include <string>
#include <vector>
#include "xlsxwriter.h"
#include "Student.cpp"
using namespace std;

class GenerateXlsx {
// 座位配置常數
const int upperRows = 4;
const int lowerRows = 3;
const int cols = 12;

// Excel 格式常數
const double COLUMN_WIDTH = 25.0;   // 欄寬
const double ROW_HEIGHT = 50.0;     // 列高
const int FONT_SIZE = 32;           // 一般字體大小
const int TITLE_FONT_SIZE = 48;     // 標題字體大小

private:
    vector<vector<string>> upperSeats;  // 上方 12x4
    vector<vector<string>> lowerSeats;  // 下方 12x3

public:
    // 生成座位表
    void generateSeating(vector<Student>& students) {
        // 初始化座位表
        upperSeats.resize(upperRows, vector<string>(cols, "空位"));
        lowerSeats.resize(lowerRows, vector<string>(cols, "空位"));
        
        int studentIndex = 0;
        int totalStudents = students.size();
        
        // 第一階段：填充上方座位的 B-K 欄 (中間 10 欄)
        for (int row = 0; row < upperRows; row++) {
            for (int col = 1; col < cols - 1; col++) {  // 跳過 A(0) 和 L(11)
                if (studentIndex < totalStudents) {
                    students[studentIndex].id = studentIndex + 1;
                    upperSeats[row][col] = students[studentIndex].name;
                    studentIndex++;
                }
            }
        }
        
        // 第二階段：填充下方座位的 B-F 和 I-K 欄 (跳過 A、G、H、L)
        for (int row = 0; row < lowerRows; row++) {
            for (int col = 1; col < cols - 1; col++) {  // 跳過 A(0) 和 L(11)
                // G 欄 (索引 6) 和 H 欄 (索引 7) 設為講台
                if (col == 6 || col == 7) {
                    lowerSeats[row][col] = "講台";
                } else {
                    if (studentIndex < totalStudents) {
                        students[studentIndex].id = studentIndex + 1;
                        lowerSeats[row][col] = students[studentIndex].name;
                        studentIndex++;
                    }
                }
            }
        }
        
        // 第三階段：最後填充 A 欄和 L 欄
        // 先填充上方的 A 欄和 L 欄
        for (int row = 0; row < upperRows; row++) {
            // A 欄 (索引 0)
            if (studentIndex < totalStudents) {
                students[studentIndex].id = studentIndex + 1;
                upperSeats[row][0] = students[studentIndex].name;
                studentIndex++;
            }
        }
        
        for (int row = 0; row < upperRows; row++) {
            // L 欄 (索引 11)
            if (studentIndex < totalStudents) {
                students[studentIndex].id = studentIndex + 1;
                upperSeats[row][11] = students[studentIndex].name;
                studentIndex++;
            }
        }
        
        // 再填充下方的 A 欄和 L 欄
        for (int row = 0; row < lowerRows; row++) {
            // A 欄 (索引 0)
            if (studentIndex < totalStudents) {
                students[studentIndex].id = studentIndex + 1;
                lowerSeats[row][0] = students[studentIndex].name;
                studentIndex++;
            }
        }
        
        for (int row = 0; row < lowerRows; row++) {
            // L 欄 (索引 11)
            if (studentIndex < totalStudents) {
                students[studentIndex].id = studentIndex + 1;
                lowerSeats[row][11] = students[studentIndex].name;
                studentIndex++;
            }
        }
    }
    
    // 顯示座位表
    void display() {
        cout << "\n========== 座位表 ==========" << endl;
        cout << "\n【上方區域 - 12欄x4列】" << endl;
        cout << "講台方向" << endl;
        cout << string(80, '-') << endl;
        
        for (int row = 0; row < upperRows; row++) {
            for (int col = 0; col < cols; col++) {
                cout << upperSeats[row][col] << "\t";
            }
            cout << endl;
        }
        
        cout << "\n" << string(80, '=') << endl;
        cout << "【走道】" << endl;
        cout << string(80, '=') << endl;
        
        cout << "\n【下方區域 - 12欄x3列】" << endl;
        for (int row = 0; row < lowerRows; row++) {
            for (int col = 0; col < cols; col++) {
                cout << lowerSeats[row][col] << "\t";
            }
            cout << endl;
        }
        cout << "\n" << string(80, '-') << endl;
    }
    
    // 儲存為 XLSX 檔案
    void saveToXLSX(string filename) {
        // 建立 workbook 和 worksheet
        lxw_workbook *workbook = workbook_new(filename.c_str());
        lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "座位表");
        
        if (!workbook) {
            cout << "無法建立 Excel 檔案!" << endl;
            return;
        }
        
        // 設定格式
        lxw_format *header_format = workbook_add_format(workbook);
        format_set_bold(header_format);
        format_set_font_size(header_format, FONT_SIZE);
        format_set_align(header_format, LXW_ALIGN_CENTER);
        format_set_bg_color(header_format, 0xD3D3D3);  // 淺灰色背景
        format_set_border(header_format, LXW_BORDER_THIN);
        
        lxw_format *title_format = workbook_add_format(workbook);
        format_set_bold(title_format);
        format_set_font_size(title_format, TITLE_FONT_SIZE);
        format_set_align(title_format, LXW_ALIGN_CENTER);
        
        lxw_format *cell_format = workbook_add_format(workbook);
        format_set_font_size(cell_format, FONT_SIZE);
        format_set_align(cell_format, LXW_ALIGN_CENTER);
        format_set_align(cell_format, LXW_ALIGN_VERTICAL_CENTER);
        format_set_border(cell_format, LXW_BORDER_THIN);
        
        lxw_format *aisle_format = workbook_add_format(workbook);
        format_set_bold(aisle_format);
        format_set_font_size(aisle_format, FONT_SIZE);
        format_set_align(aisle_format, LXW_ALIGN_CENTER);
        format_set_bg_color(aisle_format, 0xFFFF99);  // 淺黃色背景
        
        lxw_format *podium_format = workbook_add_format(workbook);
        format_set_bold(podium_format);
        format_set_font_size(podium_format, FONT_SIZE);
        format_set_align(podium_format, LXW_ALIGN_CENTER);
        format_set_align(podium_format, LXW_ALIGN_VERTICAL_CENTER);
        format_set_bg_color(podium_format, 0xFFD700);
        format_set_border(podium_format, LXW_BORDER_THIN);
        
        // 設定所有欄位為相同寬度
        worksheet_set_column(worksheet, 0, cols-1, COLUMN_WIDTH, NULL);  // 統一設定為寬度 10
        
        int currentRow = 0;
        
        // 寫入標題
        worksheet_merge_range(worksheet, currentRow, 0, currentRow, cols-1, 
                            "座位表", title_format);
        currentRow++;
        
        // 上方區域標題
        worksheet_merge_range(worksheet, currentRow, 0, currentRow, cols-1, 
                            "上方區域 (12欄 x 4列)", header_format);
        currentRow++;
        
        // 寫入上方座位
        for (int row = 0; row < upperRows; row++) {
            worksheet_set_row(worksheet, currentRow, ROW_HEIGHT, NULL);
            for (int col = 0; col < cols; col++) {
                worksheet_write_string(worksheet, currentRow, col, 
                                     upperSeats[row][col].c_str(), cell_format);
            }
            currentRow++;
        }
        
        currentRow++; 
        
        // 走道
        worksheet_merge_range(worksheet, currentRow, 0, currentRow, cols-1, 
                            "========== 走道 ==========", aisle_format);
        currentRow++;
        currentRow++; 
        
        // 下方區域標題
        worksheet_merge_range(worksheet, currentRow, 0, currentRow, cols-1, 
                            "下方區域 (12欄 x 3列)", header_format);
        currentRow++;
        
        // 寫入下方座位
        for (int row = 0; row < lowerRows; row++) {
            worksheet_set_row(worksheet, currentRow, ROW_HEIGHT, NULL);
            for (int col = 0; col < cols; col++) {
                // G 欄 (索引 6) 和 H 欄 (索引 7) 使用講台格式
                if (col == 6 || col == 7) {
                    worksheet_write_string(worksheet, currentRow, col, 
                                         lowerSeats[row][col].c_str(), podium_format);
                } else {
                    worksheet_write_string(worksheet, currentRow, col, 
                                         lowerSeats[row][col].c_str(), cell_format);
                }
            }
            currentRow++;
        }
        
        // 儲存並關閉
        lxw_error error = workbook_close(workbook);
        
        if (error) {
            cout << "儲存檔案時發生錯誤: " << lxw_strerror(error) << endl;
        } else {
            cout << "\n✓ 座位表已成功儲存至: " << filename << endl;
            cout << "  可直接用 Excel 開啟查看！" << endl;
        }
    }
};