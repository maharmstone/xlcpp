#include "xlcpp.h"
#include <archive.h>
#include <archive_entry.h>
#include <libxml/xmlwriter.h>

using namespace std;

namespace xlcpp {

sheet& workbook::add_sheet(const string& name) {
    return *sheets.emplace(sheets.end(), name, sheets.size() + 1);
}

string sheet::xml() const {
    int rc;
    string ret;
    xmlBufferPtr buf;
    xmlTextWriterPtr writer;

    buf = xmlBufferCreate();
    if (!buf)
        throw runtime_error("xmlBufferCreate failed");

    try {
        writer = xmlNewTextWriterMemory(buf, 0);
        if (!writer)
            throw runtime_error("xmlNewTextWriterMemory failed");

        try {
            rc = xmlTextWriterStartDocument(writer, nullptr, "UTF-8", nullptr);
            if (rc < 0)
                throw runtime_error("xmlTextWriterStartDocument failed (error " + to_string(rc) + ")");

            rc = xmlTextWriterStartElement(writer, BAD_CAST "test");
            if (rc < 0)
                throw runtime_error("xmlTextWriterStartElement failed (error " + to_string(rc) + ")");

            rc = xmlTextWriterWriteString(writer, BAD_CAST "he&<llo");
            if (rc < 0)
                throw runtime_error("xmlTextWriterWriteString failed (error " + to_string(rc) + ")");

            rc = xmlTextWriterEndElement(writer);
            if (rc < 0)
                throw runtime_error("xmlTextWriterEndElement failed (error " + to_string(rc) + ")");

            rc = xmlTextWriterEndDocument(writer);
            if (rc < 0)
                throw runtime_error("xmlTextWriterEndDocument failed (error " + to_string(rc) + ")");
        } catch (...) {
            xmlFreeTextWriter(writer);
            throw;
        }

        xmlFreeTextWriter(writer);
    } catch (...) {
        xmlBufferFree(buf);
        throw;
    }

    ret = (char*)buf->content;

    xmlBufferFree(buf);

    return ret;
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
