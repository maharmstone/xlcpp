#pragma once

#include <filesystem>
#include <vector>
#include <list>
#include <variant>
#include <unordered_map>

// FIXME - remove these from public headers
#include <archive.h>
#include <archive_entry.h>

// FIXME - private implementations

namespace xlcpp {

struct shared_string {
    unsigned int num;
};

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

    friend cell;

    std::list<sheet> sheets;
    std::unordered_map<std::string, shared_string> shared_strings;
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

class cell;

class row {
public:
    row(sheet& s, unsigned int num) : parent(s), num(num) { }
    cell& add_cell(int val);
    cell& add_cell(const std::string& val);
    cell& add_cell(double val);

    sheet& parent;

private:
    friend sheet;

    unsigned int num;
    std::list<cell> cells;
};

class cell {
public:
    cell(row& r, unsigned int num, int val) : parent(r), num(num), val(val) { }
    cell(row& r, unsigned int num, const std::string& val);
    cell(row& r, unsigned int num, double val) : parent(r), num(num), val(val) { }

    row& parent;

private:
    friend sheet;

    unsigned int num;
    std::variant<int, shared_string, double> val;
};

};
