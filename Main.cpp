#include <iostream>
#include <string>
#include <vector>
#include "Student.cpp"

using namespace std;
int main() {
    vector<Student> students;
    vector<string> test{"1", "2", "3"};

    for (auto t:test) {
        students.push_back(Student(t, 0));
    }

    for (auto t:test) {
        
    }
    return 0;
}