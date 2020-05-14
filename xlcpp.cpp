#include "xlcpp.h"
#include <archive.h>
#include <archive_entry.h>

using namespace std;

namespace xlcpp {

sheet& workbook::add_sheet(const string& name) {
    return *sheets.emplace(sheets.end(), name, sheets.size() + 1);
}

string sheet::xml() const {
    // FIXME
    return "<test>Lemon curry?</test>";
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
