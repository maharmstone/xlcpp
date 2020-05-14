#include "xlcpp.h"
#include <archive.h>
#include <archive_entry.h>
#include <libxml/xmlwriter.h>

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

    void start_element(const string& tag) {
        int rc = xmlTextWriterStartElement(writer, BAD_CAST tag.c_str());
        if (rc < 0)
            throw runtime_error("xmlTextWriterStartElement failed (error " + to_string(rc) + ")");
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

    writer.start_element("test");

    writer.text("he&<llo");

    writer.end_element();

    writer.end_document();

    return writer.dump();
}

void sheet::write(struct archive* a) const {
    struct archive_entry* entry;
    string data = xml();

    entry = archive_entry_new();
    archive_entry_set_pathname(entry, ("xl/worksheets/sheet" + to_string(num) + ".txt").c_str()); // FIXME
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

    // FIXME - miscellaneous files in zip

    for (const auto& sh : sheets) {
        sh.write(a);
    }

    archive_write_close(a);
    archive_write_free(a);
}

}
