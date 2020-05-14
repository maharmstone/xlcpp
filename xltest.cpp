#include <xlcpp.h>
#include <iostream>

using namespace std;

static void main2() {
    xlcpp::workbook wb;

    auto sh = wb.add_sheet("Sheet1");
    wb.save("out.xlsx");
}

int main() {
    try {
        main2();
    } catch (const exception& e) {
        cerr << "Error: " << e.what() << endl;
        return 1;
    }

    return 0;
}
