#include <iostream>
#include <string>
#include <vector>
#include "Student.cpp"
#include "ReadTxt.cpp"
#include "RandomSort.cpp"
using namespace std;

int main() {
    vector<Student> students;
    
    // 抓取 student.txt 資料
    ReadTxt rd;
    rd.getTxtData("student.txt", students);
    
    // 亂數排序 students
    RandomSort rs;
    rs.randomSort(students);

    for (auto s:students)
        cout << s.name << endl;

    return 0;
}