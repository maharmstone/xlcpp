#pragma once

#include "xlcpp.h"
#include <list>
#include <variant>
#include <unordered_map>
#include <unordered_set>
#include <functional>
#include <optional>
#include <archive.h>
#include <archive_entry.h>
#include <libxml/xmlwriter.h>
#include <libxml/xmlreader.h>

namespace xlcpp {

struct shared_string {
    unsigned int num;
};

typedef struct {
    std::string content_type;
    std::string data;
} file;

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
    workbook_pimpl() = default;
    workbook_pimpl(const std::filesystem::path& fn);
#ifdef _WIN32
    ~workbook_pimpl();
#endif
    sheet& add_sheet(const std::string& name, bool visible);
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
    la_ssize_t write_callback(struct archive* a, const void* buffer, size_t length) const;
    void parse_workbook(const std::string& fn, const std::string_view& data,
                        const std::unordered_map<std::string, file>& files);
    void load_sheet(const std::string& name, const std::string& data, bool visible);
    void load_shared_strings2(const std::string_view& sv);
    void load_shared_strings(const std::unordered_map<std::string, file>& files);
    void load_styles(const std::unordered_map<std::string, file>& files);
    void load_styles2(const std::string_view& sv);
    std::string find_number_format(unsigned int num);

#ifdef _WIN32
    void rename(const std::filesystem::path& fn) const;
#endif

    template<class... Args>
    const style* find_style(Args&&... args) {
        auto ret = styles.emplace(args...);

        if (ret.second)
            ret.first->num = styles.size() - 1;

        return &(*ret.first);
    }

    std::list<sheet> sheets;
    std::unordered_map<std::string, shared_string> shared_strings;
    std::vector<std::string> shared_strings2;
    std::unordered_set<style, style_hash> styles;
    std::unordered_map<unsigned int, std::string> number_formats;
    std::vector<std::optional<unsigned int>> cell_styles;
    bool date1904 = false;

    mutable std::string buf;

#ifdef _WIN32
    HANDLE h = INVALID_HANDLE_VALUE;
    uint8_t readbuf[1048576];
#endif
};

class sheet_pimpl {
public:
    sheet_pimpl(workbook_pimpl& wb, const std::string& name, unsigned int num, bool visible) : parent(wb), name(name), num(num), visible(visible) { }

    void write(struct archive* a) const;
    std::string xml() const;
    row& add_row();

    workbook_pimpl& parent;
    std::string name;
    unsigned int num;
    bool visible;
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
    std::variant<int64_t, shared_string, double, std::chrono::year_month_day, std::chrono::seconds, datetime, bool, std::nullptr_t, std::string> val;
    std::string number_format;
};

};

class xml_writer {
public:
    xml_writer();
    ~xml_writer();
    std::string dump() const;
    void start_document();
    void end_document();
    void start_element(const std::string& tag, const std::unordered_map<std::string, std::string>& namespaces = {});
    void end_element();
    void text(const std::string& s);
    void attribute(const std::string& name, const std::string& value);

private:
    xmlBufferPtr buf;
    xmlTextWriterPtr writer;
};

class xml_reader {
public:
    xml_reader(const std::string_view& sv);
    ~xml_reader();
    bool read();
    int node_type() const;
    std::string name() const;
    bool has_attributes() const;
    bool get_attribute(unsigned int i, std::string& name, std::string& ns, std::string& value);
    bool is_empty() const;
    void attributes_loop(const std::function<bool(const std::string&, const std::string&, const std::string&)>& func);
    std::string namespace_uri() const;
    std::string local_name() const;
    std::string value() const;

private:
    xmlParserInputBufferPtr buf;
    xmlTextReaderPtr reader;
};
