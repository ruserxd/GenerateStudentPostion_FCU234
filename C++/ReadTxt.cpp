#include <fstream>
#include <iostream>
#include <sstream>
#include <string>
#include "Student.cpp"
using namespace std;

class ReadTxt {
    public:
        void getTxtData(string file_name, vector<Student>& student_name) {
            ifstream file(file_name);

            if (!file.is_open()) {
                cout << "無法開啟檔案!" << endl;
            }

            string name = "";
            while (file >> name) {
                student_name.emplace_back(Student(name, 0));
            }

            file.close();
        }
};