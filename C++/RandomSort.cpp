#include <iostream>
#include <algorithm>
#include <string>
#include <vector>
#include <random>
#include "Student.cpp"

class RandomSort {
public:
    void randomSort(vector<Student>& students) {
        unsigned seed = std::chrono::system_clock::now().time_since_epoch().count();
        shuffle(begin(students), end(students), default_random_engine(seed));
    }
};