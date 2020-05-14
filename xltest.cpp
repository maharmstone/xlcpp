#include <xlcpp.h>

int main() {
    xlcpp::workbook wb;

    auto sh = wb.add_sheet("Sheet1");
    wb.save("out.xlsx");

    return 0;
}
