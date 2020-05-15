#pragma once

#include <filesystem>
#include <vector>
#include <list>
#include <variant>
#include <unordered_map>
#include <unordered_set>

// FIXME - remove these from public headers
#include <archive.h>
#include <archive_entry.h>

// FIXME - private implementations

namespace xlcpp {

struct shared_string {
    unsigned int num;
};

class style {
public:
    style(const std::string& number_format) : number_format(number_format) { }

    std::string number_format;
    mutable unsigned int num;
    mutable unsigned int number_format_num;
};

class style_hash {
public:
    size_t operator()(const style& s) const {
        return std::hash<std::string>{}(s.number_format);
    }
};

bool operator==(const style& lhs, const style& rhs) noexcept;

class sheet;
class cell;

class workbook {
public:
    sheet& add_sheet(const std::string& name);
    void save(const std::filesystem::path& fn) const;

private:
    void write_workbook_xml(struct archive* a) const;
    void write_content_types_xml(struct archive* a) const;
    void write_rels(struct archive* a) const;
    void write_workbook_rels(struct archive* a) const;
    shared_string get_shared_string(const std::string& s);
    void write_shared_strings(struct archive* a) const;
    void write_styles(struct archive* a) const;

    template<class... Args>
    const style* find_style(Args&&... args) {
        auto ret = styles.emplace(args...);

        if (ret.second)
            ret.first->num = styles.size() - 1;

        return &(*ret.first);
    }

    friend cell;

    std::list<sheet> sheets;
    std::unordered_map<std::string, shared_string> shared_strings;
    std::unordered_set<style, style_hash> styles;
};

class row;

class sheet {
public:
    sheet(workbook& wb, const std::string& name, unsigned int num) : parent(wb), name(name), num(num) { }
    row& add_row();

    workbook& parent;

private:
    void write(struct archive* a) const;
    std::string xml() const;

    friend workbook;

    std::string name;
    unsigned int num;
    std::list<row> rows;
};

class date {
public:
    date(unsigned int year, unsigned int month, unsigned int day) : year(year), month(month), day(day) { }

    unsigned int to_number() const;

    unsigned int year, month, day;
};

class time {
public:
    time(unsigned int hour, unsigned int minute, unsigned int second) : hour(hour), minute(minute), second(second) { }

    double to_number() const;

    unsigned int hour, minute, second;
};

class cell;

class row {
public:
    row(sheet& s, unsigned int num) : parent(s), num(num) { }

    template<typename T>
    cell& add_cell(const T& val) {
        return *cells.emplace(cells.end(), *this, cells.size() + 1, val);
    }

    sheet& parent;

private:
    friend sheet;

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
    void set_number_format(const std::string& fmt);

    row& parent;

private:
    friend sheet;

    const style* sty;

    unsigned int num;
    std::variant<int, shared_string, double, date, time> val;
};

};
