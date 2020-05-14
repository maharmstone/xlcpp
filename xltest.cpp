#include <xlcpp.h>
#include <iostream>

using namespace std;

static void main2() {
    xlcpp::workbook wb;

    auto& sh = wb.add_sheet("Sheet1");

    auto& row = sh.add_row();

    row.add_cell(1);
    row.add_cell(2);
    row.add_cell(3);

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
