#include "xlcpp-pimpl.h"

using namespace std;

xml_writer::xml_writer() {
    xmlInitParser();

    buf = xmlBufferCreate();
    if (!buf)
        throw runtime_error("xmlBufferCreate failed");

    writer = xmlNewTextWriterMemory(buf, 0);
    if (!writer) {
        xmlBufferFree(buf);
        throw runtime_error("xmlNewTextWriterMemory failed");
    }
}

xml_writer::~xml_writer() {
    xmlFreeTextWriter(writer);
    xmlBufferFree(buf);
}

string xml_writer::dump() const {
    return (char*)buf->content;
}

void xml_writer::start_document() {
    int rc = xmlTextWriterStartDocument(writer, nullptr, "UTF-8", nullptr);
    if (rc < 0)
        throw runtime_error("xmlTextWriterStartDocument failed (error " + to_string(rc) + ")");
}

void xml_writer::end_document() {
    int rc = xmlTextWriterEndDocument(writer);
    if (rc < 0)
        throw runtime_error("xmlTextWriterEndDocument failed (error " + to_string(rc) + ")");
}

void xml_writer::start_element(const string& tag, const unordered_map<string, string>& namespaces) {
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

void xml_writer::end_element() {
    int rc = xmlTextWriterEndElement(writer);
    if (rc < 0)
        throw runtime_error("xmlTextWriterEndElement failed (error " + to_string(rc) + ")");
}

void xml_writer::text(const string& s) {
    int rc = xmlTextWriterWriteString(writer, BAD_CAST s.c_str());
    if (rc < 0)
        throw runtime_error("xmlTextWriterWriteString failed (error " + to_string(rc) + ")");
}

void xml_writer::attribute(const string& name, const string& value) {
    int rc = xmlTextWriterWriteAttribute(writer, BAD_CAST name.c_str(), BAD_CAST value.c_str());
    if (rc < 0)
        throw runtime_error("xmlTextWriterWriteAttribute failed (error " + to_string(rc) + ")");
}
