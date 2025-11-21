#include <iostream>
#include <string>
#include <vector>
#include "Student.cpp"
#include "ReadTxt.cpp"
#include "RandomSort.cpp"
#include "GenerateXlsx.cpp"
using namespace std;

int main() {
    vector<Student> students;
    
    // 抓取 student.txt 資料
    ReadTxt rd;
    rd.getTxtData("student.txt", students);
    
    // 亂數排序 students
    RandomSort rs;
    rs.randomSort(students);

    // 生成 excel (xlsx)
    GenerateXlsx gx;
    gx.generateSeating(students);
    gx.display();
    gx.saveToXLSX("座位表.xlsx");
    return 0;
}