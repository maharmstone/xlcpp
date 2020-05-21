#pragma once

#include <filesystem>
#include <list>
#include <chrono>
#include <variant>

// FIXME - private implementations

namespace xlcpp {

struct shared_string {
    unsigned int num;
};

class font {
public:
    font(const std::string& font_name, unsigned int font_size, bool bold) : font_name(font_name), font_size(font_size), bold(bold) { }

    std::string font_name;
    unsigned int font_size;
    bool bold;
};

class font_hash {
public:
    size_t operator()(const font& f) const {
        return std::hash<std::string>{}(f.font_name) |
            (std::hash<unsigned int>{}(f.font_size) << 1) |
            (std::hash<bool>{}(f.bold) << 2);
    }
};

bool operator==(const font& lhs, const font& rhs) noexcept;

class style {
public:
    style(const std::string& number_format, const std::string& font, unsigned int font_size, bool bold = false) :
        number_format(number_format), font(font, font_size, bold) { }

    void set_font(const std::string& font_name, unsigned int font_size, bool bold);
    void set_number_format(const std::string& fmt);

    std::string number_format;
    xlcpp::font font;

    mutable unsigned int num;
    mutable unsigned int number_format_num;
};

class style_hash {
public:
    size_t operator()(const style& s) const {
        return std::hash<std::string>{}(s.number_format) |
            (font_hash{}(s.font) << 1);
    }
};

bool operator==(const style& lhs, const style& rhs) noexcept;

class sheet;
class cell;

class workbook_pimpl;

class workbook {
public:
    workbook();
    ~workbook();
    sheet& add_sheet(const std::string& name);
    void save(const std::filesystem::path& fn) const;

    class workbook_pimpl* impl;
};

class row;

class sheet {
public:
    sheet(workbook_pimpl& wb, const std::string& name, unsigned int num);
    ~sheet();
    row& add_row();

    class sheet_pimpl* impl;
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

class cell;

class row {
public:
    row(sheet_pimpl& s, unsigned int num) : parent(s), num(num) { }

    template<typename T>
    cell& add_cell(const T& val) {
        return *cells.emplace(cells.end(), *this, cells.size() + 1, val);
    }

    sheet_pimpl& parent;

private:
    friend sheet_pimpl;

    unsigned int num;
    std::list<cell> cells;
};

class cell {
public:
    cell(row& r, unsigned int num, int val);
    cell(row& r, unsigned int num, const std::string& val);
    cell(row& r, unsigned int num, double val);
    cell(row& r, unsigned int num, const date& val);
    cell(row& r, unsigned int num, const time& val);
    cell(row& r, unsigned int num, const datetime& val);
    cell(row& r, unsigned int num, const std::chrono::system_clock::time_point& val) : cell(r, num, datetime{val}) { }
    void set_number_format(const std::string& fmt);
    void set_font(const std::string& name, unsigned int size, bool bold = false);

    row& parent;

private:
    friend sheet_pimpl;

    const style* sty;

    unsigned int num;
    std::variant<int, shared_string, double, date, time, datetime> val;
};

};
