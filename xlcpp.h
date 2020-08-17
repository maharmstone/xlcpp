#pragma once

#include <filesystem>
#include <chrono>
#include <string>
#include <list>

#ifdef _WIN32

#ifdef XLCPP_EXPORT
#define XLCPP __declspec(dllexport)
#elif !defined(XLCPP_STATIC)
#define XLCPP __declspec(dllimport)
#else
#define XLCPP
#endif

#else

#ifdef XLCPP_EXPORT
#define XLCPP __attribute__ ((visibility ("default")))
#elif !defined(XLCPP_STATIC)
#define XLCPP __attribute__ ((dllimport))
#else
#define XLCPP
#endif

#endif

namespace xlcpp {

class workbook_pimpl;
class sheet;

class XLCPP workbook {
public:
    workbook();
    workbook(const std::filesystem::path& fn);
    ~workbook();
    sheet& add_sheet(const std::string& name);
    void save(const std::filesystem::path& fn) const;
    std::string data() const;
    const std::list<sheet>& sheets() const;

    workbook_pimpl* impl;
};

class sheet_pimpl;
class row;

class XLCPP sheet {
public:
    sheet(workbook_pimpl& wb, const std::string& name, unsigned int num);
    ~sheet();
    row& add_row();
    std::string name() const;
    const std::list<row>& rows() const;

    sheet_pimpl* impl;
};

class XLCPP date {
public:
    date(unsigned int year, unsigned int month, unsigned int day) : year(year), month(month), day(day) { }
    date(time_t tt);

    unsigned int to_number() const;
    void from_number(unsigned int num);

    unsigned int year, month, day;
};

class XLCPP time {
public:
    time(unsigned int hour, unsigned int minute, unsigned int second) : hour(hour), minute(minute), second(second) { }
    time(time_t tt);

    double to_number() const;

    unsigned int hour, minute, second;
};

class XLCPP datetime {
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

class XLCPP cell {
public:
    cell(row_pimpl& r, unsigned int num, nullptr_t);
    cell(row_pimpl& r, unsigned int num, int64_t val);
    cell(row_pimpl& r, unsigned int num, const std::string& val);
    cell(row_pimpl& r, unsigned int num, double val);
    cell(row_pimpl& r, unsigned int num, const date& val);
    cell(row_pimpl& r, unsigned int num, const time& val);
    cell(row_pimpl& r, unsigned int num, const datetime& val);
    cell(row_pimpl& r, unsigned int num, const std::chrono::system_clock::time_point& val);
    cell(row_pimpl& r, unsigned int num, bool val);

    template<typename T>
    cell(row_pimpl& r, unsigned int num, T* t) = delete;

    void set_number_format(const std::string& fmt);
    void set_font(const std::string& name, unsigned int size, bool bold = false);
    std::string get_number_format() const;

    cell_pimpl* impl;
};

XLCPP std::ostream& operator<<(std::ostream& os, const cell& c);

class XLCPP row {
public:
    row(sheet_pimpl& s, unsigned int num);
    ~row();

    cell& add_cell(int64_t val);
    cell& add_cell(const std::string& val);
    cell& add_cell(double val);
    cell& add_cell(const date& val);
    cell& add_cell(const time& val);
    cell& add_cell(const datetime& val);
    cell& add_cell(const std::chrono::system_clock::time_point& val);
    cell& add_cell(bool val);
    cell& add_cell(nullptr_t);

    cell& add_cell(const char* val) {
        return add_cell(std::string(val));
    }

    cell& add_cell(char* val) {
        return add_cell(std::string(val));
    }

    cell& add_cell(int val) {
        return add_cell((int64_t)val);
    }

    template<typename T>
    cell& add_cell(T* val) = delete;

    const std::list<cell>& cells() const;

    row_pimpl* impl;
};

};
