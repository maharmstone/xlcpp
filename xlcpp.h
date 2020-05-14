#pragma once

#include <filesystem>
#include <vector>
#include <string>

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

    std::vector<sheet> sheets;
};

class sheet {
public:
    sheet(const std::string& name, unsigned int num) : name(name), num(num) { }

private:
    void write(struct archive* a) const;
    std::string xml() const;

    friend workbook;

    std::string name;
    unsigned int num;
};

};
