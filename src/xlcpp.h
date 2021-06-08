#pragma once

#include <filesystem>
#include <chrono>
#include <string>
#include <list>
#include <variant>

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
    sheet& add_sheet(const std::string& name, bool visible = true);
    void save(const std::filesystem::path& fn) const;
    std::string data() const;
    const std::list<sheet>& sheets() const;
#ifdef _WIN32
    void rename(const std::filesystem::path& fn) const;
#endif

    workbook_pimpl* impl;
};

class sheet_pimpl;
class row;

class XLCPP sheet {
public:
    sheet(workbook_pimpl& wb, const std::string& name, unsigned int num, bool visible = true);
    ~sheet();
    row& add_row();
    std::string name() const;
    bool visible() const;
    const std::list<row>& rows() const;

    sheet_pimpl* impl;
};

class XLCPP datetime {
public:
    constexpr datetime(std::chrono::year year, std::chrono::month month, std::chrono::day day, std::chrono::hours hour, std::chrono::minutes minute, std::chrono::seconds second) :
        d(year, month, day), t(hour + minute + second) { }
    constexpr datetime(std::chrono::year year, std::chrono::month month, std::chrono::day day, std::chrono::seconds t) :
        d(year, month, day), t(t) { }

    template<typename T>
    constexpr datetime(const std::chrono::time_point<T>& chr) :
        d(std::chrono::floor<std::chrono::days>(chr)),
        t(std::chrono::floor<std::chrono::seconds>(chr - std::chrono::floor<std::chrono::days>(chr)).count()) {
    }

    double to_number(bool date1904 = false) const;

    std::chrono::year_month_day d;
    std::chrono::seconds t;
};

class row_pimpl;
class cell_pimpl;

using cell_t = std::variant<int64_t, std::string, double, std::chrono::year_month_day, std::chrono::seconds, datetime, bool, std::nullptr_t>;

class XLCPP cell {
public:
    cell(row_pimpl& r, unsigned int num, std::nullptr_t);
    cell(row_pimpl& r, unsigned int num, int64_t val);
    cell(row_pimpl& r, unsigned int num, const std::string& val);
    cell(row_pimpl& r, unsigned int num, double val);
    cell(row_pimpl& r, unsigned int num, const std::chrono::year_month_day& val);
    cell(row_pimpl& r, unsigned int num, const std::chrono::seconds& val);
    cell(row_pimpl& r, unsigned int num, const datetime& val);
    cell(row_pimpl& r, unsigned int num, const std::chrono::system_clock::time_point& val);
    cell(row_pimpl& r, unsigned int num, bool val);

    template<typename T>
    cell(row_pimpl& r, unsigned int num, T* t) = delete;

    void set_number_format(const std::string& fmt);
    void set_font(const std::string& name, unsigned int size, bool bold = false);
    std::string get_number_format() const;
    cell_t value() const;

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
    cell& add_cell(const std::chrono::year_month_day& val);
    cell& add_cell(const std::chrono::seconds& val);
    cell& add_cell(const datetime& val);
    cell& add_cell(const std::chrono::system_clock::time_point& val);
    cell& add_cell(bool val);
    cell& add_cell(std::nullptr_t);

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
