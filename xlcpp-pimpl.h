#pragma once

#include "xlcpp.h"
#include <list>
#include <unordered_map>
#include <unordered_set>
#include <archive.h>
#include <archive_entry.h>

namespace xlcpp {

class workbook_pimpl {
public:
    sheet& add_sheet(const std::string& name);
    void save(const std::filesystem::path& fn) const;

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

    std::list<sheet> sheets;
    std::unordered_map<std::string, shared_string> shared_strings;
    std::unordered_set<style, style_hash> styles;
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

};
