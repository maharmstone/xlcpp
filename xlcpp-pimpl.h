#pragma once

#include "xlcpp.h"
#include <list>
#include <variant>
#include <unordered_map>
#include <unordered_set>
#include <archive.h>
#include <archive_entry.h>

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

class workbook_pimpl {
public:
    sheet& add_sheet(const std::string& name);
    void save(const std::filesystem::path& fn) const;
    std::string data() const;

    void write_workbook_xml(struct archive* a) const;
    void write_content_types_xml(struct archive* a) const;
    void write_rels(struct archive* a) const;
    void write_workbook_rels(struct archive* a) const;
    shared_string get_shared_string(const std::string& s);
    void write_shared_strings(struct archive* a) const;
    void write_styles(struct archive* a) const;
    void write_archive(struct archive* a) const;
    ssize_t write_callback(struct archive* a, const void* buffer, size_t length) const;

    template<class... Args>
    const style* find_style(Args&&... args) {
        auto ret = styles.emplace(args...);

        if (ret.second)
            ret.first->num = styles.size() - 1;

        return &(*ret.first);
    }

    std::list<sheet> sheets;
    std::unordered_map<std::string, shared_string> shared_strings;
    std::unordered_set<style, style_hash> styles;

    mutable std::string buf;
};

class sheet_pimpl {
public:
    sheet_pimpl(workbook_pimpl& wb, const std::string& name, unsigned int num) : parent(wb), name(name), num(num) { }

    void write(struct archive* a) const;
    std::string xml() const;
    row& add_row();

    workbook_pimpl& parent;
    std::string name;
    unsigned int num;
    std::list<row> rows;
};

class row_pimpl {
public:
    row_pimpl(sheet_pimpl& s, unsigned int num) : parent(s), num(num) { }

    template<typename T>
    cell& add_cell(const T& val) {
        return *cells.emplace(cells.end(), *this, cells.size() + 1, val);
    }

    sheet_pimpl& parent;
    unsigned int num;
    std::list<cell> cells;
};

class cell_pimpl {
public:
    template<typename T>
    cell_pimpl(row_pimpl& r, unsigned int num, const T& t);

    cell_pimpl(row_pimpl& r, unsigned int num, const std::chrono::system_clock::time_point& val) : cell_pimpl(r, num, datetime{val}) { }

    void set_number_format(const std::string& fmt);
    void set_font(const std::string& name, unsigned int size, bool bold = false);

    row_pimpl& parent;

    const style* sty;

    unsigned int num;
    std::variant<int64_t, shared_string, double, date, time, datetime> val;
};

};
