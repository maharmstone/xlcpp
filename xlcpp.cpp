#include "xlcpp.h"
#include "xlcpp-pimpl.h"
#include <archive.h>
#include <archive_entry.h>
#include <libxml/xmlwriter.h>
#include <vector>

using namespace std;

class xml_writer {
public:
    xml_writer() {
        buf = xmlBufferCreate();
        if (!buf)
            throw runtime_error("xmlBufferCreate failed");

        writer = xmlNewTextWriterMemory(buf, 0);
        if (!writer) {
            xmlBufferFree(buf);
            throw runtime_error("xmlNewTextWriterMemory failed");
        }
    }

    ~xml_writer() {
        xmlFreeTextWriter(writer);
        xmlBufferFree(buf);
    }

    string dump() const {
        return (char*)buf->content;
    }

    void start_document() {
        int rc = xmlTextWriterStartDocument(writer, nullptr, "UTF-8", nullptr);
        if (rc < 0)
            throw runtime_error("xmlTextWriterStartDocument failed (error " + to_string(rc) + ")");
    }

    void end_document() {
        int rc = xmlTextWriterEndDocument(writer);
        if (rc < 0)
            throw runtime_error("xmlTextWriterEndDocument failed (error " + to_string(rc) + ")");
    }

    void start_element(const string& tag, const unordered_map<string, string>& namespaces = {}) {
        int rc = xmlTextWriterStartElement(writer, BAD_CAST tag.c_str());
        if (rc < 0)
            throw runtime_error("xmlTextWriterStartElement failed (error " + to_string(rc) + ")");

        for (const auto& ns : namespaces) {
            string att = ns.first.empty() ? "xmlns" : ("xmlns:" + ns.first);

            rc = xmlTextWriterWriteAttribute(writer, BAD_CAST att.c_str(), BAD_CAST ns.second.c_str());
            if (rc < 0)
                throw runtime_error("xmlTextWriterWriteAttribute failed (error " + to_string(rc) + ")");
        }
    }

    void end_element() {
        int rc = xmlTextWriterEndElement(writer);
        if (rc < 0)
            throw runtime_error("xmlTextWriterEndElement failed (error " + to_string(rc) + ")");
    }

    void text(const string& s) {
        int rc = xmlTextWriterWriteString(writer, BAD_CAST s.c_str());
        if (rc < 0)
            throw runtime_error("xmlTextWriterWriteString failed (error " + to_string(rc) + ")");
    }

    void attribute(const string& name, const string& value) {
        int rc = xmlTextWriterWriteAttribute(writer, BAD_CAST name.c_str(), BAD_CAST value.c_str());
        if (rc < 0)
            throw runtime_error("xmlTextWriterWriteAttribute failed (error " + to_string(rc) + ")");
    }

private:
    xmlBufferPtr buf;
    xmlTextWriterPtr writer;
};

namespace xlcpp {

sheet& workbook_pimpl::add_sheet(const string& name) {
    return *sheets.emplace(sheets.end(), *this, name, sheets.size() + 1);
}

sheet& workbook::add_sheet(const string& name) {
    return impl->add_sheet(name);
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

string sheet_pimpl::xml() const {
    xml_writer writer;

    writer.start_document();

    writer.start_element("worksheet", {{"", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}});

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
                writer.text(to_string(get<date>(c.impl->val).to_number()));
                writer.end_element();
            } else if (holds_alternative<time>(c.impl->val)) {
                writer.attribute("t", "n"); // number

                writer.start_element("v");
                writer.text(to_string(get<time>(c.impl->val).to_number()));
                writer.end_element();
            } else if (holds_alternative<datetime>(c.impl->val)) {
                writer.attribute("t", "n"); // number

                writer.start_element("v");
                writer.text(to_string(get<datetime>(c.impl->val).to_number()));
                writer.end_element();
            } else if (holds_alternative<bool>(c.impl->val)) {
                writer.attribute("t", "b"); // bool

                writer.start_element("v");
                writer.text(to_string(get<bool>(c.impl->val)));
                writer.end_element();
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

        writer.start_element("workbook", {{"", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}, {"r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}});

        writer.start_element("workbookPr");
        writer.attribute("date1904", "true");
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

        writer.start_element("Types", {{"", "http://schemas.openxmlformats.org/package/2006/content-types"}});

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

        writer.start_element("sst", {{"", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}});
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
        unordered_map<font, unsigned int, font_hash> fonts;
        vector<const font*> fonts2;

        writer.start_document();

        writer.start_element("styleSheet");
        writer.attribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

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
            writer.attribute("numFmtId", to_string(i));
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

        sty.resize(styles.size());

        for (const auto& s : styles) {
            sty[s.num] = &s;
        }

        writer.start_element("cellStyleXfs");
        writer.attribute("count", to_string(styles.size()));

        for (const auto& s : sty) {
            writer.start_element("xf");
            writer.attribute("numFmtId", to_string(number_formats[s->number_format]));
            writer.attribute("fontId", to_string(fonts[s->font]));
            writer.end_element();
        }

        writer.end_element();

        {
            unsigned int style_id = 0;

            writer.start_element("cellXfs");
            writer.attribute("count", to_string(styles.size()));

            for (const auto& s : sty) {
                writer.start_element("xf");
                writer.attribute("numFmtId", to_string(number_formats[s->number_format]));
                writer.attribute("fontId", to_string(fonts[s->font]));
                writer.attribute("applyFont", "1"); // FIXME - "0" if not specified explicitly?
                writer.attribute("xfId", to_string(style_id));
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

ssize_t workbook_pimpl::write_callback(struct archive* a, const void* buffer, size_t length) const {
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

unsigned int date::to_number() const {
    int m2 = ((int)month - 14) / 12;
    long long n;

    n = (1461 * ((int)year + 4800 + m2)) / 4;
    n += (367 * ((int)month - 2 - (12 * m2))) / 12;
    n -= (3 * (((int)year + 4900 + m2)/100)) / 4;
    n += day;
    n -= 2448556;

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

double datetime::to_number() const {
    return (double)d.to_number() + t.to_number();
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

time::time(time_t tt) {
    tm local_tm = *localtime(&tt);

    hour = local_tm.tm_hour;
    minute = local_tm.tm_min;
    second = local_tm.tm_sec;
}

workbook::workbook() {
    impl = new workbook_pimpl;
}

workbook::~workbook() {
    delete impl;
}

sheet::sheet(workbook_pimpl& wb, const std::string& name, unsigned int num) {
    impl = new sheet_pimpl(wb, name, num);
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

}
