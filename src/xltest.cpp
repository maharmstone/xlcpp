#include <xlcpp.h>
#include <iostream>
#include <chrono>

using namespace std;

static void main2() {
    xlcpp::workbook wb;

    auto& sh = wb.add_sheet("Sheet1");

    auto& row = sh.add_row();
    row.add_cell(1);
    row.add_cell(2);
    row.add_cell(3);

    auto& row2 = sh.add_row();

    auto& c = row2.add_cell("hello");
    c.set_font("Comic Sans MS", 12, true);

    row2.add_cell(42.1);
    row2.add_cell(true);

    auto& c2 = row2.add_cell(chrono::year_month_day{1998y, chrono::July, 5d});
    c2.set_number_format("YYYY-MM-DD");
    c2.set_font("Comic Sans MS", 12);

    row2.add_cell(chrono::hours{12} + chrono::minutes{34} + chrono::seconds{56});

    row2.add_cell(chrono::system_clock::now());

    wb.save("out.xlsx");

//     auto res = wb.data();
//     cout << res;
}

static void read_test(const filesystem::path& fn) {
    xlcpp::workbook wb(fn);

    for (const auto& sh : wb.sheets()) {
        unsigned int row_num = 0;

        cout << "Sheet: " << sh.name() << endl;

        for (const auto& r : sh.rows()) {
            cout << "Row " << row_num << endl;

            for (const auto& c : r.cells()) {
                cout << "Cell: " << c << " (" << c.get_number_format() << ")" << endl;
            }

            row_num++;
        }
    }
}

int main() {
    try {
        main2();
        read_test("out.xlsx");
    } catch (const exception& e) {
        cerr << "Error: " << e.what() << endl;
        return 1;
    }

    return 0;
}
