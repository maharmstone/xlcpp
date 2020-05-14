#pragma once

#include <filesystem>
#include <vector>
#include <list>

// FIXME - remove these from public headers
#include <archive.h>
#include <archive_entry.h>

// FIXME - private implementations

namespace xlcpp {

class sheet;

class workbook {
public:
    sheet& add_sheet(const std::string& name);
    void save(const std::filesystem::path& fn) const;

private:
    void write_workbook_xml(struct archive* a) const;
    void write_content_types_xml(struct archive* a) const;
    void write_rels(struct archive* a) const;
    void write_workbook_rels(struct archive* a) const;

    std::list<sheet> sheets;
};

class row;

class sheet {
public:
    sheet(const std::string& name, unsigned int num) : name(name), num(num) { }
    row& add_row();

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
    row(unsigned int num) : num(num) { }
    cell& add_cell(int val);

private:
    friend sheet;

    unsigned int num;
    std::list<cell> cells;
};

class cell {
public:
    cell(unsigned int num, int val) : num(num), val(val) { }

private:
    friend sheet;

    unsigned int num;
    int val;
};

};
