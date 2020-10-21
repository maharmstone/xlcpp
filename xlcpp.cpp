#include "xlcpp.h"
#include "xlcpp-pimpl.h"
#include <archive.h>
#include <archive_entry.h>
#include <vector>
#include <array>

#define FMT_HEADER_ONLY
#include <fmt/format.h>

#define BLOCK_SIZE 20480

using namespace std;

static const string NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
static const string NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
static const string NS_PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
static const string NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";

#define NUMFMT_OFFSET 165

static const array builtin_styles = {
    pair{ 0, "General" },
    pair{ 1, "0" },
    pair{ 2, "0.00" },
    pair{ 3, "#,##0" },
    pair{ 4, "#,##0.00" },
    pair{ 9, "0%" },
    pair{ 10, "0.00%" },
    pair{ 11, "0.00E+00" },
    pair{ 12, "# ?/?" },
    pair{ 13, "# ??/??" },
    pair{ 14, "mm-dd-yy" },
    pair{ 15, "d-mmm-yy" },
    pair{ 16, "d-mmm" },
    pair{ 17, "mmm-yy" },
    pair{ 18, "h:mm AM/PM" },
    pair{ 19, "h:mm:ss AM/PM" },
    pair{ 20, "h:mm" },
    pair{ 21, "h:mm:ss" },
    pair{ 22, "m/d/yy h:mm" },
    pair{ 37, "#,##0 ;(#,##0)" },
    pair{ 38, "#,##0 ;[Red](#,##0)" },
    pair{ 39, "#,##0.00;(#,##0.00)" },
    pair{ 40, "#,##0.00;[Red](#,##0.00)" },
    pair{ 45, "mm:ss" },
    pair{ 46, "[h]:mm:ss" },
    pair{ 47, "mmss.0" },
    pair{ 48, "##0.0E+0" },
    pair{ 49, "@" },
};

namespace xlcpp {

sheet& workbook_pimpl::add_sheet(const string& name, bool visible) {
    return *sheets.emplace(sheets.end(), *this, name, sheets.size() + 1, visible);
}

sheet& workbook::add_sheet(const string& name, bool visible) {
    return impl->add_sheet(name, visible);
}

static string make_reference(unsigned int row, unsigned int col) {
    char colstr[4];

    col--;

    if (col < 26) {
        colstr[0] = col + 'A';
        colstr[1] = 0;
    } else if (col < 702) {
        colstr[0] = ((col / 26) - 1) + 'A';
        colstr[1] = (col % 26) + 'A';
        colstr[2] = 0;
    } else // FIXME - support three-letter columns
        throw runtime_error("Column " + to_string(col) + " too large.");

    return string(colstr) + to_string(row);
}

static void resolve_reference(const string_view& sv, unsigned int& row, unsigned int& col) {
    if (sv.length() >= 2 && sv[0] >= 'A' && sv[0] <= 'Z' && sv[1] >= '0' && sv[1] <= '9') {
        col = sv[0] - 'A';
        row = stoi(string(sv.data() + 1, sv.length() - 1)) - 1;
        return;
    } else if (sv.length() >= 3 && sv[0] >= 'A' && sv[0] <= 'Z' && sv[1] >= 'A' && sv[1] <= 'Z' && sv[2] >= '0' && sv[2] <= '9') {
        col = ((sv[0] - 'A' + 1) * 26) + sv[1] - 'A';
        row = stoi(string(sv.data() + 2, sv.length() - 2)) - 1;
        return;
    }

    // FIXME - support three-letter columns

    throw runtime_error("Malformed reference \"" + string(sv) + "\".");
}

string sheet_pimpl::xml() const {
    xml_writer writer;

    writer.start_document();

    writer.start_element("worksheet", {{"", NS_SPREADSHEET}});

    writer.start_element("sheetData");

    for (const auto& r : rows) {
        writer.start_element("row");

        writer.attribute("r", to_string(r.impl->num));
        writer.attribute("customFormat", "false");
        writer.attribute("ht", "12.8");
        writer.attribute("hidden", "false");
        writer.attribute("customHeight", "false");
        writer.attribute("outlineLevel", "0");
        writer.attribute("collapsed", "false");

        for (const auto& c : r.impl->cells) {
            writer.start_element("c");

            writer.attribute("r", make_reference(r.impl->num, c.impl->num));
            writer.attribute("s", to_string(c.impl->sty->num));

            if (holds_alternative<int64_t>(c.impl->val)) {
                writer.attribute("t", "n"); // number

                writer.start_element("v");
                writer.text(to_string(get<int64_t>(c.impl->val)));
                writer.end_element();
            } else if (holds_alternative<shared_string>(c.impl->val)) {
                writer.attribute("t", "s"); // shared string

                writer.start_element("v");
                writer.text(to_string(get<shared_string>(c.impl->val).num));
                writer.end_element();
            } else if (holds_alternative<double>(c.impl->val)) {
                writer.attribute("t", "n"); // number

                writer.start_element("v");
                writer.text(to_string(get<double>(c.impl->val)));
                writer.end_element();
            } else if (holds_alternative<date>(c.impl->val)) {
                writer.attribute("t", "n"); // number

                writer.start_element("v");
                writer.text(to_string(get<date>(c.impl->val).to_number(parent.date1904)));
                writer.end_element();
            } else if (holds_alternative<time>(c.impl->val)) {
                writer.attribute("t", "n"); // number

                writer.start_element("v");
                writer.text(to_string(get<time>(c.impl->val).to_number()));
                writer.end_element();
            } else if (holds_alternative<datetime>(c.impl->val)) {
                writer.attribute("t", "n"); // number

                writer.start_element("v");
                writer.text(to_string(get<datetime>(c.impl->val).to_number(parent.date1904)));
                writer.end_element();
            } else if (holds_alternative<bool>(c.impl->val)) {
                writer.attribute("t", "b"); // bool

                writer.start_element("v");
                writer.text(to_string(get<bool>(c.impl->val)));
                writer.end_element();
            } else if (holds_alternative<nullptr_t>(c.impl->val)) {
                // nop
            } else
                throw runtime_error("Unknown type for cell.");

            writer.end_element();
        }

        writer.end_element();
    }

    writer.end_element();

    writer.end_element();

    writer.end_document();

    return writer.dump();
}

void sheet_pimpl::write(struct archive* a) const {
    struct archive_entry* entry;
    string data = xml();

    entry = archive_entry_new();
    archive_entry_set_pathname(entry, ("xl/worksheets/sheet" + to_string(num) + ".xml").c_str());
    archive_entry_set_size(entry, data.length());
    archive_entry_set_filetype(entry, AE_IFREG);
    archive_entry_set_perm(entry, 0644);
    archive_write_header(a, entry);
    archive_write_data(a, data.data(), data.length());
    // FIXME - set date?
    archive_entry_free(entry);
}

void workbook_pimpl::write_workbook_xml(struct archive* a) const {
    struct archive_entry* entry;
    string data;

    {
        xml_writer writer;

        writer.start_document();

        writer.start_element("workbook", {{"", NS_SPREADSHEET}, {"r", NS_RELATIONSHIPS}});

        writer.start_element("workbookPr");
        writer.attribute("date1904", date1904 ? "true" : "false");
        writer.end_element();

        writer.start_element("sheets");

        for (const auto& sh : sheets) {
            writer.start_element("sheet");
            writer.attribute("name", sh.impl->name);
            writer.attribute("sheetId", to_string(sh.impl->num));
            writer.attribute("state", "visible");
            writer.attribute("r:id", "rId" + to_string(sh.impl->num));
            writer.end_element();
        }

        writer.end_element();

        writer.end_element();

        writer.end_document();

        data = move(writer.dump());
    }

    entry = archive_entry_new();
    archive_entry_set_pathname(entry, "xl/workbook.xml");
    archive_entry_set_size(entry, data.length());
    archive_entry_set_filetype(entry, AE_IFREG);
    archive_entry_set_perm(entry, 0644);
    archive_write_header(a, entry);
    archive_write_data(a, data.data(), data.length());
    // FIXME - set date?
    archive_entry_free(entry);
}

void workbook_pimpl::write_content_types_xml(struct archive* a) const {
    struct archive_entry* entry;
    string data;

    {
        xml_writer writer;

        writer.start_document();

        writer.start_element("Types", {{"", NS_CONTENT_TYPES}});

        writer.start_element("Override");
        writer.attribute("PartName", "/xl/workbook.xml");
        writer.attribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        writer.end_element();

        for (const auto& sh : sheets) {
            writer.start_element("Override");
            writer.attribute("PartName", "/xl/worksheets/sheet" + to_string(sh.impl->num) + ".xml");
            writer.attribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            writer.end_element();
        }

        writer.start_element("Override");
        writer.attribute("PartName", "/_rels/.rels");
        writer.attribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
        writer.end_element();

        writer.start_element("Override");
        writer.attribute("PartName", "/xl/_rels/workbook.xml.rels");
        writer.attribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
        writer.end_element();

        if (!shared_strings.empty()) {
            writer.start_element("Override");
            writer.attribute("PartName", "/xl/sharedStrings.xml");
            writer.attribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
            writer.end_element();
        }

        if (!styles.empty()) {
            writer.start_element("Override");
            writer.attribute("PartName", "/xl/styles.xml");
            writer.attribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            writer.end_element();
        }

        writer.end_element();

        writer.end_document();

        data = move(writer.dump());
    }

    entry = archive_entry_new();
    archive_entry_set_pathname(entry, "[Content_Types].xml");
    archive_entry_set_size(entry, data.length());
    archive_entry_set_filetype(entry, AE_IFREG);
    archive_entry_set_perm(entry, 0644);
    archive_write_header(a, entry);
    archive_write_data(a, data.data(), data.length());
    // FIXME - set date?
    archive_entry_free(entry);
}

void workbook_pimpl::write_rels(struct archive* a) const {
    struct archive_entry* entry;
    string data;

    {
        xml_writer writer;

        writer.start_document();

        writer.start_element("Relationships", {{"", "http://schemas.openxmlformats.org/package/2006/relationships"}});

        writer.start_element("Relationship");
        writer.attribute("Id", "rId1");
        writer.attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        writer.attribute("Target", "/xl/workbook.xml");
        writer.end_element();

        writer.end_element();

        writer.end_document();

        data = move(writer.dump());
    }

    entry = archive_entry_new();
    archive_entry_set_pathname(entry, "_rels/.rels");
    archive_entry_set_size(entry, data.length());
    archive_entry_set_filetype(entry, AE_IFREG);
    archive_entry_set_perm(entry, 0644);
    archive_write_header(a, entry);
    archive_write_data(a, data.data(), data.length());
    // FIXME - set date?
    archive_entry_free(entry);
}

void workbook_pimpl::write_workbook_rels(struct archive* a) const {
    struct archive_entry* entry;
    string data;

    {
        unsigned int num = 1;
        xml_writer writer;

        writer.start_document();

        writer.start_element("Relationships", {{"", "http://schemas.openxmlformats.org/package/2006/relationships"}});

        for (const auto& sh : sheets) {
            writer.start_element("Relationship");
            writer.attribute("Id", "rId" + to_string(num));
            writer.attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            writer.attribute("Target", "worksheets/sheet" + to_string(sh.impl->num) + ".xml");
            writer.end_element();
            num++;
        }

        if (!shared_strings.empty()) {
            writer.start_element("Relationship");
            writer.attribute("Id", "rId" + to_string(num));
            writer.attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
            writer.attribute("Target", "sharedStrings.xml");
            writer.end_element();
            num++;
        }

        if (!styles.empty()) {
            writer.start_element("Relationship");
            writer.attribute("Id", "rId" + to_string(num));
            writer.attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            writer.attribute("Target", "styles.xml");
            writer.end_element();
        }

        writer.end_element();

        writer.end_document();

        data = move(writer.dump());
    }

    entry = archive_entry_new();
    archive_entry_set_pathname(entry, "xl/_rels/workbook.xml.rels");
    archive_entry_set_size(entry, data.length());
    archive_entry_set_filetype(entry, AE_IFREG);
    archive_entry_set_perm(entry, 0644);
    archive_write_header(a, entry);
    archive_write_data(a, data.data(), data.length());
    // FIXME - set date?
    archive_entry_free(entry);
}

void workbook_pimpl::write_shared_strings(struct archive* a) const {
    struct archive_entry* entry;
    string data;

    {
        xml_writer writer;

        writer.start_document();

        writer.start_element("sst", {{"", NS_SPREADSHEET}});
        writer.attribute("count", to_string(shared_strings.size()));
        writer.attribute("uniqueCount", to_string(shared_strings.size()));

        vector<string> strings;

        strings.resize(shared_strings.size());

        for (const auto& ss : shared_strings) {
            strings[ss.second.num] = ss.first;
        }

        for (const auto& s : strings) {
            writer.start_element("si");

            writer.start_element("t");
            writer.attribute("xml:space", "preserve");
            writer.text(s);
            writer.end_element();

            writer.end_element();
        }

        writer.end_element();

        writer.end_document();

        data = move(writer.dump());
    }

    entry = archive_entry_new();
    archive_entry_set_pathname(entry, "xl/sharedStrings.xml");
    archive_entry_set_size(entry, data.length());
    archive_entry_set_filetype(entry, AE_IFREG);
    archive_entry_set_perm(entry, 0644);
    archive_write_header(a, entry);
    archive_write_data(a, data.data(), data.length());
    // FIXME - set date?
    archive_entry_free(entry);
}

void workbook_pimpl::write_styles(struct archive* a) const {
    struct archive_entry* entry;
    string data;

    {
        xml_writer writer;
        vector<const style*> sty;
        vector<unsigned int> number_format_nums;
        unordered_map<string, unsigned int> number_formats;
        vector<const string*> number_formats2;
        font default_font("Arial", 10, false);
        unordered_map<font, unsigned int, font_hash> fonts;
        vector<const font*> fonts2;

        writer.start_document();

        writer.start_element("styleSheet");
        writer.attribute("xmlns", NS_SPREADSHEET);

        fonts[default_font] = 0;
        fonts2.emplace_back(&default_font);

        for (const auto& s : styles) {
            if (number_formats.find(s.number_format) == number_formats.end()) {
                number_formats[s.number_format] = number_formats.size();
                number_formats2.emplace_back(&s.number_format);
            }

            if (fonts.find(s.font) == fonts.end()) {
                fonts[s.font] = fonts.size();
                fonts2.emplace_back(&s.font);
            }
        }

        writer.start_element("numFmts");
        writer.attribute("count", to_string(number_formats.size()));

        for (unsigned int i = 0; i < number_formats.size(); i++) {
            writer.start_element("numFmt");
            writer.attribute("numFmtId", to_string(i + NUMFMT_OFFSET));
            writer.attribute("formatCode", *number_formats2[i]);
            writer.end_element();
        }

        writer.end_element();

        writer.start_element("fonts");
        writer.attribute("count", to_string(fonts.size()));

        for (unsigned int i = 0; i < fonts.size(); i++) {
            writer.start_element("font");

            if (fonts2[i]->bold) {
                writer.start_element("b");
                writer.attribute("val", "true");
                writer.end_element();
            }

            writer.start_element("sz");
            writer.attribute("val", to_string(fonts2[i]->font_size));
            writer.end_element();

            writer.start_element("name");
            writer.attribute("val", fonts2[i]->font_name);
            writer.end_element();

            writer.end_element();
        }

        writer.end_element();

        writer.start_element("fills");
        writer.attribute("count", "2");

        writer.start_element("fill");
        writer.start_element("patternFill");
        writer.attribute("patternType", "none");
        writer.end_element();
        writer.end_element();

        writer.start_element("fill");
        writer.start_element("patternFill");
        writer.attribute("patternType", "gray125");
        writer.end_element();
        writer.end_element();

        writer.end_element();

        writer.start_element("borders");
        writer.attribute("count", "1");

        writer.start_element("border");
        writer.start_element("left"); writer.end_element();
        writer.start_element("right"); writer.end_element();
        writer.start_element("top"); writer.end_element();
        writer.start_element("bottom"); writer.end_element();
        writer.start_element("diagonal"); writer.end_element();
        writer.end_element();

        writer.end_element();

        sty.resize(styles.size());

        for (const auto& s : styles) {
            sty[s.num] = &s;
        }

        {
            unsigned int style_id = 0;

            writer.start_element("cellXfs");
            writer.attribute("count", to_string(styles.size()));

            for (const auto& s : sty) {
                writer.start_element("xf");
                writer.attribute("numFmtId", to_string(number_formats[s->number_format] + NUMFMT_OFFSET));
                writer.attribute("fontId", to_string(fonts[s->font]));
                writer.attribute("fillId", "0");
                writer.attribute("borderId", "0");
                writer.attribute("applyFont", "true");
                writer.attribute("applyNumberFormat", "true");
                writer.end_element();

                style_id++;
            }

            writer.end_element();
        }

        writer.end_element();

        writer.end_document();

        data = move(writer.dump());
    }

    entry = archive_entry_new();
    archive_entry_set_pathname(entry, "xl/styles.xml");
    archive_entry_set_size(entry, data.length());
    archive_entry_set_filetype(entry, AE_IFREG);
    archive_entry_set_perm(entry, 0644);
    archive_write_header(a, entry);
    archive_write_data(a, data.data(), data.length());
    // FIXME - set date?
    archive_entry_free(entry);
}

void workbook_pimpl::write_archive(struct archive* a) const {
    for (const auto& sh : sheets) {
        sh.impl->write(a);
    }

    write_workbook_xml(a);
    write_content_types_xml(a);
    write_rels(a);
    write_workbook_rels(a);

    if (!shared_strings.empty())
        write_shared_strings(a);

    if (!styles.empty())
        write_styles(a);
}

void workbook_pimpl::save(const filesystem::path& fn) const {
    struct archive* a;

    a = archive_write_new();
    archive_write_set_format_zip(a);

    archive_write_open_filename(a, fn.u8string().c_str());

    write_archive(a);

    archive_write_close(a);
    archive_write_free(a);
}

void workbook::save(const filesystem::path& fn) const {
    impl->save(fn);
}

static int archive_dummy_callback(struct archive* a, void* client_data) {
    return ARCHIVE_OK;
}

la_ssize_t workbook_pimpl::write_callback(struct archive* a, const void* buffer, size_t length) const {
    buf.append((const char*)buffer, length);

    return length;
}

string workbook_pimpl::data() const {
    struct archive* a;
    int ret;

    buf.clear();

    a = archive_write_new();

    try {
        archive_write_set_format_zip(a);
        archive_write_set_bytes_in_last_block(a, 1);

        ret = archive_write_open(a, (void*)this, archive_dummy_callback,
                                [](struct archive* a, void* client_data, const void* buffer, size_t length) {
                                    auto wb = (const workbook_pimpl*)client_data;

                                    return wb->write_callback(a, buffer, length);
                                },
                                archive_dummy_callback);
        if (ret != ARCHIVE_OK)
            throw runtime_error("archive_write_open returned " + to_string(ret) + ".");

        write_archive(a);

        archive_write_close(a);
    } catch (...) {
        archive_write_free(a);
        throw;
    }

    archive_write_free(a);

    return buf;
}

string workbook::data() const {
    return impl->data();
}

row& sheet_pimpl::add_row() {
    return *rows.emplace(rows.end(), *this, rows.size() + 1);
}

row& sheet::add_row() {
    return impl->add_row();
}

shared_string workbook_pimpl::get_shared_string(const string& s) {
    shared_string ss;

    if (shared_strings.count(s) != 0)
        return shared_strings.at(s);

    ss.num = shared_strings.size();

    shared_strings.emplace(s, ss);

    return ss;
}

template<typename T>
cell_pimpl::cell_pimpl(row_pimpl& r, unsigned int num, const T& t) : parent(r), num(num), val(t) {
    if constexpr (std::is_same_v<T, date>)
        sty = parent.parent.parent.find_style(style("dd/mm/yy", "Arial", 10)); // FIXME - localization
    else if constexpr (std::is_same_v<T, time>)
        sty = parent.parent.parent.find_style(style("HH:MM:SS", "Arial", 10)); // FIXME - localization
    else if constexpr (std::is_same_v<T, datetime>)
        sty = parent.parent.parent.find_style(style("dd/mm/yy HH:MM:SS", "Arial", 10)); // FIXME - localization
    else
        sty = parent.parent.parent.find_style(style("General", "Arial", 10));
}

template<>
cell_pimpl::cell_pimpl(row_pimpl& r, unsigned int num, const string& t) : cell_pimpl(r, num, r.parent.parent.get_shared_string(t)) {
}

cell::cell(row_pimpl& r, unsigned int num, int64_t val) {
    impl = new cell_pimpl(r, num, val);
}

cell::cell(row_pimpl& r, unsigned int num, const string& val) {
    impl = new cell_pimpl(r, num, val);
}

cell::cell(row_pimpl& r, unsigned int num, double val) {
    impl = new cell_pimpl(r, num, val);
}

cell::cell(row_pimpl& r, unsigned int num, const date& val) {
    impl = new cell_pimpl(r, num, val);
}

cell::cell(row_pimpl& r, unsigned int num, const time& val) {
    impl = new cell_pimpl(r, num, val);
}

cell::cell(row_pimpl& r, unsigned int num, const datetime& val) {
    impl = new cell_pimpl(r, num, val);
}

cell::cell(row_pimpl& r, unsigned int num, const chrono::system_clock::time_point& val) {
    impl = new cell_pimpl(r, num, val);
}

cell::cell(row_pimpl& r, unsigned int num, bool val) {
    impl = new cell_pimpl(r, num, val);
}

cell::cell(row_pimpl& r, unsigned int num, nullptr_t) {
    impl = new cell_pimpl(r, num, nullptr);
}

unsigned int date::to_number(bool date1904) const {
    int m2 = ((int)month - 14) / 12;
    long long n;

    n = (1461 * ((int)year + 4800 + m2)) / 4;
    n += (367 * ((int)month - 2 - (12 * m2))) / 12;
    n -= (3 * (((int)year + 4900 + m2)/100)) / 4;
    n += day;
    n -= 2447094;

    if (date1904)
        n -= 1462;

    return n;
}

bool operator==(const font& lhs, const font& rhs) noexcept {
    return lhs.font_name == rhs.font_name &&
        lhs.font_size == rhs.font_size &&
       ((lhs.bold && rhs.bold) || (!lhs.bold && !rhs.bold));
}

bool operator==(const style& lhs, const style& rhs) noexcept {
    return lhs.number_format == rhs.number_format &&
        lhs.font == rhs.font;
}

double time::to_number() const {
    return (double)((hour * 3600) + (minute * 60) + second) / 86400.0;
}

double datetime::to_number(bool date1904) const {
    return (double)d.to_number(date1904) + t.to_number();
}

void style::set_font(const std::string& font_name, unsigned int font_size, bool bold) {
    this->font = xlcpp::font(font_name, font_size, bold);
}

void style::set_number_format(const std::string& fmt) {
    number_format = fmt;
}

void cell_pimpl::set_font(const std::string& name, unsigned int size, bool bold) {
    auto sty2 = *sty;

    sty2.set_font(name, size, bold);

    sty = parent.parent.parent.find_style(sty2);
}

void cell::set_font(const std::string& name, unsigned int size, bool bold) {
    impl->set_font(name, size, bold);
}

void cell_pimpl::set_number_format(const std::string& fmt) {
    auto sty2 = *sty;

    sty2.set_number_format(fmt);

    sty = parent.parent.parent.find_style(sty2);
}

void cell::set_number_format(const std::string& fmt) {
    impl->set_number_format(fmt);
}

date::date(time_t tt) {
    tm local_tm = *localtime(&tt);

    year = local_tm.tm_year + 1900;
    month = local_tm.tm_mon + 1;
    day = local_tm.tm_mday;
}

void date::from_number(unsigned int num, bool date1904) {
    unsigned int J = num + 2415019;
    unsigned int f, e, g, h;

    if (date1904)
        J += 1462;

    f = J;
    f *= 4;
    f += 274277;
    f /= 146097;
    f *= 3;
    f /= 4;
    f += J;
    f += 1363;

    e = (f * 4) + 3;

    g = e % 1461;
    g /= 4;

    h = (5 * g) + 2;

    day = h % 153;
    day /= 5;
    day++;

    month = h;
    month /= 153;
    month += 2;
    month %= 12;
    month++;

    year = 14 - month;
    year /= 12;
    year -= 4716;
    year += e / 1461;
}

time::time(time_t tt) {
    tm local_tm = *localtime(&tt);

    hour = local_tm.tm_hour;
    minute = local_tm.tm_min;
    second = local_tm.tm_sec;
}

workbook::workbook() {
    impl = new workbook_pimpl;
    impl->date1904 = true;
}

static void parse_content_types(const string& ct, unordered_map<string, file>& files) {
    xml_reader r(ct);
    unsigned int depth = 0;
    unordered_map<string, string> defs, over;

    while (r.read()) {
        unsigned int next_depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT && !r.is_empty())
            next_depth = depth + 1;
        else if (r.node_type() == XML_READER_TYPE_END_ELEMENT)
            next_depth = depth - 1;
        else
            next_depth = depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT) {
            if (depth == 0) {
                if (r.local_name() != "Types" || r.namespace_uri() != NS_CONTENT_TYPES)
                    throw runtime_error("Root tag name was not \"Types\".");
            } else if (depth == 1) {
                if (r.local_name() == "Default" && r.namespace_uri() == NS_CONTENT_TYPES) {
                    string ext, ct;

                    r.attributes_loop([&](const string& name, const string& ns, const string& value) {
                        if (name == "Extension" && ns.empty())
                            ext = value;
                        else if (name == "ContentType" && ns.empty())
                            ct = value;

                        return true;
                    });

                    defs[ext] = ct;
                } else if (r.local_name() == "Override" && r.namespace_uri() == NS_CONTENT_TYPES) {
                    string part, ct;

                    r.attributes_loop([&](const string& name, const string& ns, const string& value) {
                        if (name == "PartName" && ns.empty())
                            part = value;
                        else if (name == "ContentType" && ns.empty())
                            ct = value;

                        return true;
                    });

                    over[part] = ct;
                }
            }
        }

        depth = next_depth;
    }

    for (auto& f : files) {
        for (const auto& o : over) {
            if (o.first == "/" + f.first) {
                f.second.content_type = o.second;
                break;
            }
        }

        if (!f.second.content_type.empty())
            continue;

        auto st = f.first.rfind(".");
        string ext = st == string::npos ? "" : f.first.substr(st + 1);

        for (const auto& d : defs) {
            if (ext == d.first) {
                f.second.content_type = d.second;
                break;
            }
        }
    }
}

static const pair<const string, file>& find_workbook(const unordered_map<string, file>& files) {
    for (const auto& f : files) {
        if (f.second.content_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
            return f;
    }

    throw runtime_error("Could not find workbook XML file.");
}

static unordered_map<string, string> read_relationships(const string& fn, const unordered_map<string, file>& files) {
    filesystem::path p = fn;
    unordered_map<string, string> rels;

    p.remove_filename();
    p /= "_rels";
    p /= filesystem::path(fn).filename().u8string() + ".rels";

    auto ps = p.u8string();

    for (auto& c : ps) {
        if (c == '\\')
            c = '/';
    }

    if (files.count(ps) == 0)
        throw runtime_error("File " + ps + " not found.");

    xml_reader r(files.at(ps).data);
    unsigned int depth = 0;

    while (r.read()) {
        unsigned int next_depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT && !r.is_empty())
            next_depth = depth + 1;
        else if (r.node_type() == XML_READER_TYPE_END_ELEMENT)
            next_depth = depth - 1;
        else
            next_depth = depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT) {
            if (depth == 0) {
                if (r.local_name() != "Relationships" || r.namespace_uri() != NS_PACKAGE_RELATIONSHIPS)
                    throw runtime_error("Root tag name was not \"Relationships\".");
            } else if (depth == 1) {
                if (r.local_name() == "Relationship" && r.namespace_uri() == NS_PACKAGE_RELATIONSHIPS) {
                    string id, target;

                    r.attributes_loop([&](const string& name, const string& ns, const string& value) {
                        if (name == "Id" && ns.empty())
                            id = value;
                        else if (name == "Target" && ns.empty())
                            target = value;

                        return true;
                    });

                    if (!id.empty() && !target.empty())
                        rels[id] = target;
                }
            }
        }

        depth = next_depth;
    }

    return rels;
}

string workbook_pimpl::find_number_format(unsigned int num) {
    if (num >= cell_styles.size())
        return "";

    if (!cell_styles[num].has_value())
        return "";

    unsigned int numfmtid = cell_styles[num].value();

    for (const auto& nf : number_formats) {
        if (nf.first == numfmtid)
            return nf.second;
    }

    for (const auto& bs : builtin_styles) {
        if ((unsigned int)bs.first == numfmtid)
            return bs.second;
    }

    return "";
}

static bool is_date(const string_view& sv) {
    if (sv == "General")
        return false;

    string s;

    s.reserve(sv.length());

    for (auto c : sv) {
        if (c >= 'a' && c <= 'z') {
            if (s.empty() || s.back() != c)
                s += c;
        } else if (c >= 'A' && c <= 'Z') {
            if (s.empty() || s.back() != c - 'A' + 'a')
                s += c - 'A' + 'a';
        }
    }

    static const char* patterns[] = {
        "dmy",
        "ymd",
        "mdy"
    };

    for (const auto& p : patterns) {
        if (s.find(p) != string::npos)
            return true;
    }

    return false;
}

static bool is_time(const string_view& sv) {
    if (sv == "General")
        return false;

    string s;

    s.reserve(sv.length());

    for (auto c : sv) {
        if (c >= 'a' && c <= 'z') {
            if (s.empty() || s.back() != c)
                s += c;
        } else if (c >= 'A' && c <= 'Z') {
            if (s.empty() || s.back() != c - 'A' + 'a')
                s += c - 'A' + 'a';
        }
    }

    return s.find("hm") != string::npos;
}

static bool __inline is_hex(char c) {
    return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f');
}

static uint8_t __inline from_hex(char c) {
    if (c >= '0' && c <= '9')
        return c - '0';
    else if (c >= 'a' && c <= 'f')
        return c - 'a' + 0xa;
    else if (c >= 'A' && c <= 'F')
        return c - 'A' + 0xa;

    return 0;
}

static string decode_escape_sequences(const string_view& sv) {
    bool has_underscore = false;

    if (sv.length() < 7)
        return string(sv);

    for (auto c : sv) {
        if (c == '_') {
            has_underscore = true;
            break;
        }
    }

    if (!has_underscore)
        return string(sv);

    string s;

    s.reserve(sv.length());

    for (unsigned int i = 0; i < sv.length(); i++) {
        if (i + 6 < sv.length() && sv[i] == '_' && sv[i+1] == 'x' && is_hex(sv[i+2]) && is_hex(sv[i+3]) && is_hex(sv[i+4]) && is_hex(sv[i+5]) && sv[i+6] == '_') {
            uint16_t val = (from_hex(sv[i+2]) << 12) | (from_hex(sv[i+3]) << 8) | (from_hex(sv[i+4]) << 4) | from_hex(sv[i+5]);

            if (val < 0x80)
                s += (char)val;
            else if (val < 0x800) {
                char buf[3];

                buf[0] = 0xc0 | (val >> 6);
                buf[1] = 0x80 | (val & 0x3f);
                buf[2] = 0;

                s += buf;
            } else {
                char buf[4];

                buf[0] = 0xe0 | (val >> 12);
                buf[1] = 0x80 | ((val & 0xfc0) >> 6);
                buf[2] = 0x80 | (val & 0x3f);
                buf[3] = 0;

                s += buf;
            }

            i += 6;
        } else
            s += sv[i];
    }

    return s;
}

void workbook_pimpl::load_sheet(const string& name, const string& data, bool visible) {
    auto& s = *sheets.emplace(sheets.end(), *this, name, sheets.size() + 1, visible);

    xml_reader r(data);
    unsigned int depth = 0;
    bool in_sheet_data = false;
    unsigned int last_index = 0, last_col;
    row* row = nullptr;
    string r_val, s_val, t_val, v_val;
    bool in_v = false;

    while (r.read()) {
        unsigned int next_depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT && !r.is_empty())
            next_depth = depth + 1;
        else if (r.node_type() == XML_READER_TYPE_END_ELEMENT)
            next_depth = depth - 1;
        else
            next_depth = depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT) {
            if (depth == 0) {
                if (r.local_name() != "worksheet" || r.namespace_uri() != NS_SPREADSHEET)
                    throw runtime_error("Root tag name was not \"worksheet\".");
            } else if (depth == 1 && r.local_name() == "sheetData" && r.namespace_uri() == NS_SPREADSHEET && !r.is_empty())
                in_sheet_data = true;
            else if (in_sheet_data) {
                if (r.local_name() == "row" && r.namespace_uri() == NS_SPREADSHEET) {
                    unsigned int row_index = 0;

                    r.attributes_loop([&](const string& name, const string& ns, const string& value) {
                        if (name == "r" && ns.empty()) {
                            row_index = stoi(value);
                            return false;
                        }

                        return true;
                    });

                    if (row_index == 0)
                        throw runtime_error("No row index given.");

                    if (row_index <= last_index)
                        throw runtime_error("Rows out of order.");

                    while (last_index + 1 < row_index) {
                        s.impl->rows.emplace(s.impl->rows.end(), *s.impl, s.impl->rows.size() + 1);
                        last_index++;
                    }

                    s.impl->rows.emplace(s.impl->rows.end(), *s.impl, s.impl->rows.size() + 1);

                    row = &s.impl->rows.back();

                    last_index = row_index;
                    last_col = 0;
                } else if (row && r.local_name() == "c" && r.namespace_uri() == NS_SPREADSHEET) {
                    unsigned int row_num, col_num;

                    r_val = t_val = s_val = v_val = "";

                    r.attributes_loop([&](const string& name, const string& ns, const string& value) {
                        if (name == "r" && ns.empty())
                            r_val = value;
                        else if (name == "t" && ns.empty())
                            t_val = value;
                        else if (name == "s" && ns.empty())
                            s_val = value;

                        return true;
                    });

                    if (r_val.empty())
                        throw runtime_error("Cell had no r value.");

                    resolve_reference(r_val, row_num, col_num);

                    if (row_num + 1 != last_index)
                        throw runtime_error("Cell was in wrong row.");

                    if (col_num < last_col)
                        throw runtime_error("Cells out of order.");

                    while (last_col < col_num) {
                        row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                        last_col++;
                    }

                    last_col = col_num + 1;

                    if (r.is_empty())
                        row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                } else if (row && r.local_name() == "v" && r.namespace_uri() == NS_SPREADSHEET && !r.is_empty())
                    in_v = true;
            }
        } else if (r.node_type() == XML_READER_TYPE_END_ELEMENT) {
            if (depth == 1 && r.local_name() == "sheetData" && r.namespace_uri() == NS_SPREADSHEET)
                in_sheet_data = false;
            else if (depth == 2 && r.local_name() == "row" && r.namespace_uri() == NS_SPREADSHEET)
                row = nullptr;
            else if (row && r.local_name() == "c" && r.namespace_uri() == NS_SPREADSHEET) {
                cell* c;
                string number_format;

                // FIXME - d, date
                // FIXME - inlineStr, inline string

                if (!s_val.empty())
                    number_format = find_number_format(stoi(s_val));

                if (t_val == "n" || t_val.empty()) { // number
                    bool dt = is_date(number_format);
                    bool tm = is_time(number_format);

                    // FIXME - we can optimize is_date and is_time if one of the preset number formats

                    if (dt && tm) {
                        auto d = stod(v_val);
                        auto n = (unsigned int)((d - (int)d) * 86400.0);
                        datetime dt(1970, 1, 1, n / 3600, (n % 3600) / 60, n % 60);

                        dt.d.from_number((int)d, date1904);

                        c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, dt);
                    } else if (dt) {
                        date d(1970, 1, 1);
                        int num;
                        bool valid_num = true;

                        try {
                            num = stoi(v_val);
                        } catch (...) {
                            valid_num = false;
                        }

                        if (valid_num) {
                            d.from_number(num, date1904);
                            c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, d);
                        } else
                            c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                    } else if (tm) {
                        auto n = (unsigned int)(stod(v_val) * 86400.0);

                        n %= 86400;
                        time t(n / 3600, (n % 3600) / 60, n % 60);

                        c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, t);
                    } else {
                        bool is_int = true;

                        for (const auto c : v_val) {
                            if (c != '-' && (c < '0' || c > '9')) {
                                is_int = false;
                                break;
                            }
                        }

                        if (is_int)
                            c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, (int64_t)stoll(v_val));
                        else
                            c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, stod(v_val));
                    }
                } else if (t_val == "b") // boolean
                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, stoi(v_val) != 0);
                else if (t_val == "s") { // shared string
                    shared_string ss;

                    ss.num = stoi(v_val);

                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);

                    // so we don't have to expose shared_string publicly
                    delete c->impl;
                    c->impl = new cell_pimpl(*row->impl, row->impl->cells.size(), ss);
                } else if (t_val == "e") // error
                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                else if (t_val == "str") { // string
                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);

                    if (!v_val.empty()) {
                        // so we don't have to expose shared_string publicly
                        c->impl->val = decode_escape_sequences(v_val);
                    }
                } else
                    throw runtime_error("Unhandled cell type value \"" + t_val + "\".");

                if (!s_val.empty())
                    c->impl->number_format = number_format;
            } else if (in_v && r.local_name() == "v" && r.namespace_uri() == NS_SPREADSHEET)
                in_v = false;
        } else if (r.node_type() == XML_READER_TYPE_TEXT) {
            if (in_v)
                v_val += r.value();
        }

        depth = next_depth;
    }
}

void workbook_pimpl::parse_workbook(const string& fn, const string_view& data, const unordered_map<string, file>& files) {
    xml_reader r(data);
    unsigned int depth = 0;

    struct sheet_info {
        sheet_info(const string_view& rid, const string_view& name, bool visible) :
            rid(rid), name(name), visible(visible) { }

        string rid;
        string name;
        bool visible;
    };

    vector<sheet_info> sheets_rels;

    while (r.read()) {
        unsigned int next_depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT && !r.is_empty())
            next_depth = depth + 1;
        else if (r.node_type() == XML_READER_TYPE_END_ELEMENT)
            next_depth = depth - 1;
        else
            next_depth = depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT) {
            if (depth == 0) {
                if (r.local_name() != "workbook" || r.namespace_uri() != NS_SPREADSHEET)
                    throw runtime_error("Root tag name was not \"workbook\".");
            } else if (depth == 1) {
                if (r.local_name() == "workbookPr" && r.namespace_uri() == NS_SPREADSHEET) {
                    r.attributes_loop([&](const string& name, const string& ns, const string& value) {
                        if (name == "date1904" && ns.empty())
                            date1904 = value == "true" || value == "1";

                        return true;
                    });
                }
            } else if (depth == 2) {
                if (r.local_name() == "sheet" && r.namespace_uri() == NS_SPREADSHEET) {
                    string sheet_name, rid;
                    bool visible = true;

                    r.attributes_loop([&](const string& name, const string& ns, const string& value) {
                        if (name == "name" && ns.empty())
                            sheet_name = value;
                        else if (name == "id" && ns == NS_RELATIONSHIPS)
                            rid = value;
                        else if (name == "state" && ns.empty())
                            visible = value != "hidden";

                        return true;
                    });

                    if (!sheet_name.empty() && !rid.empty())
                        sheets_rels.emplace_back(rid, sheet_name, visible);
                }
            }
        }

        depth = next_depth;
    }

    // FIXME - preserve sheet order

    auto rels = read_relationships(fn, files);

    for (const auto& sr : sheets_rels) {
        for (const auto& r : rels) {
            if (r.first == sr.rid) {
                auto name = filesystem::path(fn);

                // FIXME - can we resolve relative paths properly?

                name.remove_filename();
                name /= r.second;

                auto ns = name.u8string();

                for (auto& c : ns) {
                    if (c == '\\')
                        c = '/';
                }

                if (files.count(ns) == 0)
                    throw runtime_error("File " + ns + " not found.");

                load_sheet(sr.name, files.at(ns).data, sr.visible);
                break;
            }
        }
    }
}

void workbook_pimpl::load_shared_strings2(const string_view& sv) {
    xml_reader r(sv);
    unsigned int depth = 0;
    bool in_si = false;
    string si_val;

    while (r.read()) {
        unsigned int next_depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT && !r.is_empty())
            next_depth = depth + 1;
        else if (r.node_type() == XML_READER_TYPE_END_ELEMENT)
            next_depth = depth - 1;
        else
            next_depth = depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT) {
            if (depth == 0) {
                if (r.local_name() != "sst" || r.namespace_uri() != NS_SPREADSHEET)
                    throw runtime_error("Root tag name was not \"sst\".");
            } else if (depth == 1) {
                if (r.local_name() == "si" && r.namespace_uri() == NS_SPREADSHEET) {
                    in_si = true;
                    si_val = "";
                }
            }
        } else if (r.node_type() == XML_READER_TYPE_TEXT) {
            if (in_si)
                si_val += decode_escape_sequences(r.value());
        } else if (r.node_type() == XML_READER_TYPE_END_ELEMENT) {
            if (r.local_name() == "si" && r.namespace_uri() == NS_SPREADSHEET) {
                shared_strings2.emplace_back(si_val);
                in_si = false;
            }
        }

        depth = next_depth;
    }
}

void workbook_pimpl::load_shared_strings(const unordered_map<string, file>& files) {
    for (const auto& f : files) {
        if (f.second.content_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml") {
            load_shared_strings2(f.second.data);
            return;
        }
    }
}

void workbook_pimpl::load_styles2(const string_view& sv) {
    xml_reader r(sv);
    unsigned int depth = 0;
    bool in_numfmts = false, in_cellxfs = false;

    while (r.read()) {
        unsigned int next_depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT && !r.is_empty())
            next_depth = depth + 1;
        else if (r.node_type() == XML_READER_TYPE_END_ELEMENT)
            next_depth = depth - 1;
        else
            next_depth = depth;

        if (r.node_type() == XML_READER_TYPE_ELEMENT) {
            if (depth == 0) {
                if (r.local_name() != "styleSheet" || r.namespace_uri() != NS_SPREADSHEET)
                    throw runtime_error("Root tag name was not \"styleSheet\".");
            } else if (depth == 1) {
                if (r.local_name() == "numFmts" && r.namespace_uri() == NS_SPREADSHEET && !r.is_empty())
                    in_numfmts = true;
                else if (r.local_name() == "cellXfs" && r.namespace_uri() == NS_SPREADSHEET && !r.is_empty())
                    in_cellxfs = true;
            } else if (depth == 2) {
                if (in_numfmts && r.local_name() == "numFmt" && r.namespace_uri() == NS_SPREADSHEET) {
                    unsigned int id = 0;
                    string format_code;

                    r.attributes_loop([&](const string& name, const string& ns, const string& value) {
                        if (name == "numFmtId" && ns.empty())
                            id = stoi(value);
                        else if (name == "formatCode" && ns.empty())
                            format_code = value;

                        return true;
                    });

                    if (id != 0)
                        number_formats[id] = format_code;
                } else if (in_cellxfs && r.local_name() == "xf" && r.namespace_uri() == NS_SPREADSHEET) {
                    optional<unsigned int> numfmtid;
                    bool apply_number_format = true;

                    r.attributes_loop([&](const string& name, const string& ns, const string& value) {
                        if (name == "numFmtId" && ns.empty())
                            numfmtid = stoi(value);
                        else if (name == "applyNumberFormat" && ns.empty())
                            apply_number_format = value == "true" || value == "1";

                        return true;
                    });

                    if (!apply_number_format)
                        numfmtid = nullopt;

                    cell_styles.push_back(numfmtid);
                }
            }
        } else if (r.node_type() == XML_READER_TYPE_END_ELEMENT) {
            if (in_numfmts && r.local_name() == "numFmts" && r.namespace_uri() == NS_SPREADSHEET)
                in_numfmts = false;
            else if (in_cellxfs && r.local_name() == "cellXfs" && r.namespace_uri() == NS_SPREADSHEET)
                in_cellxfs = false;
        }

        depth = next_depth;
    }
}

void workbook_pimpl::load_styles(const unordered_map<string, file>& files) {
    for (const auto& f : files) {
        if (f.second.content_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml") {
            load_styles2(f.second.data);
            return;
        }
    }
}

workbook_pimpl::workbook_pimpl(const filesystem::path& fn) {
    struct archive* a = archive_read_new();
    struct archive_entry* entry;

    try {
        archive_read_support_format_zip(a);

        auto r = archive_read_open_filename(a, fn.u8string().c_str(), BLOCK_SIZE);

        if (r != ARCHIVE_OK)
            throw runtime_error(archive_error_string(a));

        unordered_map<string, file> files;

        while (archive_read_next_header(a, &entry) == ARCHIVE_OK) {
            if (archive_entry_filetype(entry) == AE_IFREG && archive_entry_pathname(entry)) {
                filesystem::path name = archive_entry_pathname(entry);
                auto ext = name.extension().u8string();

                for (auto& c : ext) {
                    if (c >= 'A' && c <= 'Z')
                        c = c - 'A' + 'a';
                }

                if (ext != ".xml" && ext != ".rels")
                    continue;

                string buf;
                string tmp(BLOCK_SIZE, 0);

                do {
                    auto read = archive_read_data(a, tmp.data(), BLOCK_SIZE);

                    if (read == 0)
                        break;

                    buf += tmp.substr(0, read);
                } while (true);

                files[name.u8string()].data = buf;
            }
        }

        if (files.count("[Content_Types].xml") == 0)
            throw runtime_error("[Content_Types].xml not found.");

        parse_content_types(files.at("[Content_Types].xml").data, files);

        load_shared_strings(files);

        load_styles(files);

        auto& wb = find_workbook(files);

        parse_workbook(wb.first, wb.second.data, files);

        // FIXME
    } catch (...) {
        archive_read_free(a);
        throw;
    }

    archive_read_free(a);
}

workbook::workbook(const filesystem::path& fn) {
    impl = new workbook_pimpl(fn);
}

workbook::~workbook() {
    delete impl;
}

sheet::sheet(workbook_pimpl& wb, const std::string& name, unsigned int num, bool visible) {
    impl = new sheet_pimpl(wb, name, num, visible);
}

sheet::~sheet() {
    delete impl;
}

row::row(sheet_pimpl& s, unsigned int num) {
    impl = new row_pimpl(s, num);
}

row::~row() {
    delete impl;
}

cell& row::add_cell(int64_t val) {
    return impl->add_cell(val);
}

cell& row::add_cell(const std::string& val) {
    return impl->add_cell(val);
}

cell& row::add_cell(double val) {
    return impl->add_cell(val);
}

cell& row::add_cell(const date& val) {
    return impl->add_cell(val);
}

cell& row::add_cell(const time& val) {
    return impl->add_cell(val);
}

cell& row::add_cell(const datetime& val) {
    return impl->add_cell(val);
}

cell& row::add_cell(const chrono::system_clock::time_point& val) {
    return impl->add_cell(val);
}

cell& row::add_cell(bool val) {
    return impl->add_cell(val);
}

cell& row::add_cell(nullptr_t) {
    return impl->add_cell(nullptr);
}

const list<sheet>& workbook::sheets() const {
    return impl->sheets;
}

std::string sheet::name() const {
    return impl->name;
}

bool sheet::visible() const {
    return impl->visible;
}

const std::list<row>& sheet::rows() const {
    return impl->rows;
}

const std::list<cell>& row::cells() const {
    return impl->cells;
}

std::ostream& operator<<(std::ostream& os, const cell& c) {
    if (holds_alternative<int64_t>(c.impl->val))
        os << get<int64_t>(c.impl->val);
    else if (holds_alternative<double>(c.impl->val))
        os << get<double>(c.impl->val);
    else if (holds_alternative<bool>(c.impl->val))
        os << (get<bool>(c.impl->val) ? "true" : "false");
    else if (holds_alternative<shared_string>(c.impl->val))
        os << c.impl->parent.parent.parent.shared_strings2[get<shared_string>(c.impl->val).num];
    else if (holds_alternative<date>(c.impl->val)) {
        const auto& d = get<date>(c.impl->val);

        os << fmt::format("{:04}-{:02}-{:02}", d.year, d.month, d.day);
    } else if (holds_alternative<time>(c.impl->val)) {
        const auto& t = get<time>(c.impl->val);

        os << fmt::format("{:02}:{:02}:{:02}", t.hour, t.minute, t.second);
    } else if (holds_alternative<datetime>(c.impl->val)) {
        const auto& dt = get<datetime>(c.impl->val);

        os << fmt::format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                          dt.d.year, dt.d.month, dt.d.day, dt.t.hour, dt.t.minute, dt.t.second);
    } else if (holds_alternative<nullptr_t>(c.impl->val)) {
        // nop
    } else
        os << "?";

    return os;
}

std::string cell::get_number_format() const {
    return impl->number_format;
}

variant<int64_t, string, double, date, time, datetime, bool, nullptr_t> cell::value() const {
    decltype(value()) v;

    visit([&](auto&& arg) {
        if constexpr (is_same_v<decay_t<decltype(arg)>, shared_string>)
            v = impl->parent.parent.parent.shared_strings2[get<shared_string>(impl->val).num];
        else
            v = arg;
    }, impl->val);

    return v;
}

}
