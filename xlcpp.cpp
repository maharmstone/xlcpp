#include "xlcpp.h"
#include <archive.h>
#include <archive_entry.h>
#include <libxml/xmlwriter.h>
#include <unordered_map>

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
            string att = ns.first.empty() ? "xmlns" : ("xmlns" + ns.first);

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

sheet& workbook::add_sheet(const string& name) {
    return *sheets.emplace(sheets.end(), name, sheets.size() + 1);
}

string sheet::xml() const {
    xml_writer writer;

    writer.start_document();

    writer.start_element("worksheet", {{"", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}});

    writer.start_element("sheetData");
    // FIXME
    writer.end_element();

    writer.end_element();

    writer.end_document();

    return writer.dump();
}

void sheet::write(struct archive* a) const {
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

void workbook::write_workbook_xml(struct archive* a) const {
    struct archive_entry* entry;
    string data;

    {
        xml_writer writer;

        writer.start_document();

        writer.start_element("workbook", {{"", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}});

        writer.start_element("sheets");

        for (const auto& sh : sheets) {
            writer.start_element("sheet");
            writer.attribute("name", sh.name);
            writer.attribute("sheetId", to_string(sh.num));
            writer.attribute("state", "visible");
            // FIXME - id?
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

void workbook::write_content_types_xml(struct archive* a) const {
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
            writer.attribute("PartName", "/xl/worksheets/sheet" + to_string(sh.num) + ".xml");
            writer.attribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
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

void workbook::save(const filesystem::path& fn) const {
    struct archive* a;

    a = archive_write_new();
    archive_write_set_format_zip(a);

    archive_write_open_filename(a, fn.u8string().c_str());

    // FIXME - [Content_Types].xml

    for (const auto& sh : sheets) {
        sh.write(a);
    }

    write_workbook_xml(a);
    write_content_types_xml(a);

    archive_write_close(a);
    archive_write_free(a);
}

}
