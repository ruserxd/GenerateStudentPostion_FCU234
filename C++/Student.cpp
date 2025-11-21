#include <iostream>
#include <string>
// 定義一次
#pragma once

using namespace std;

class Student {
public:
    string name;
    int id;

    Student(const string& name_, int id_) : name(name_), id(id_) {}
};