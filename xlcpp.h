#pragma once

#include <filesystem>
#include <chrono>

namespace xlcpp {

class workbook_pimpl;
class sheet;

class workbook {
public:
    workbook();
    ~workbook();
    sheet& add_sheet(const std::string& name);
    void save(const std::filesystem::path& fn) const;
    std::string data() const;

    workbook_pimpl* impl;
};

class sheet_pimpl;
class row;

class sheet {
public:
    sheet(workbook_pimpl& wb, const std::string& name, unsigned int num);
    ~sheet();
    row& add_row();

    sheet_pimpl* impl;
};

class date {
public:
    date(unsigned int year, unsigned int month, unsigned int day) : year(year), month(month), day(day) { }
    date(time_t tt);

    unsigned int to_number() const;

    unsigned int year, month, day;
};

class time {
public:
    time(unsigned int hour, unsigned int minute, unsigned int second) : hour(hour), minute(minute), second(second) { }
    time(time_t tt);

    double to_number() const;

    unsigned int hour, minute, second;
};

class datetime {
public:
    datetime(unsigned int year, unsigned int month, unsigned int day, unsigned int hour, unsigned int minute, unsigned int second):
        d(year, month, day), t(hour, minute, second) { }
    datetime(time_t tt) : d(tt), t(tt) { }
    datetime(const std::chrono::system_clock::time_point& tp) : datetime(std::chrono::system_clock::to_time_t(tp)) { }

    double to_number() const;

    date d;
    time t;
};

class row_pimpl;
class cell_pimpl;

class cell {
public:
    cell(row_pimpl& r, unsigned int num, int val);
    cell(row_pimpl& r, unsigned int num, const std::string& val);
    cell(row_pimpl& r, unsigned int num, double val);
    cell(row_pimpl& r, unsigned int num, const date& val);
    cell(row_pimpl& r, unsigned int num, const time& val);
    cell(row_pimpl& r, unsigned int num, const datetime& val);
    cell(row_pimpl& r, unsigned int num, const std::chrono::system_clock::time_point& val);

    void set_number_format(const std::string& fmt);
    void set_font(const std::string& name, unsigned int size, bool bold = false);

    cell_pimpl* impl;
};

class row {
public:
    row(sheet_pimpl& s, unsigned int num);
    ~row();

    cell& add_cell(int val);
    cell& add_cell(const std::string& val);
    cell& add_cell(double val);
    cell& add_cell(const date& val);
    cell& add_cell(const time& val);
    cell& add_cell(const datetime& val);
    cell& add_cell(const std::chrono::system_clock::time_point& val);

    row_pimpl* impl;
};

};
