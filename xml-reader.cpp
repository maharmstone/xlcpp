#include "xlcpp-pimpl.h"

using namespace std;

xml_reader::xml_reader(const string_view& sv) {
    buf = xmlParserInputBufferCreateMem(sv.data(), sv.length(), XML_CHAR_ENCODING_UTF8);
    if (!buf)
        throw runtime_error("xmlParserInputBufferCreateMem failed");

    reader = xmlNewTextReader(buf, nullptr);
    if (!reader) {
        xmlFreeParserInputBuffer(buf);
        throw runtime_error("xmlNewTextReader failed");
    }
}

xml_reader::~xml_reader() {
    xmlFreeTextReader(reader);
    xmlFreeParserInputBuffer(buf);
}

bool xml_reader::read() {
    auto ret = xmlTextReaderRead(reader);

    return ret == 1;
}

int xml_reader::node_type() const {
    return xmlTextReaderNodeType(reader);
}

string xml_reader::name() const {
    auto xc = xmlTextReaderConstName(reader);

    if (!xc)
        return "";

    return (char*)xc;
}

bool xml_reader::has_attributes() const {
    return xmlTextReaderHasAttributes(reader) != 0;
}

bool xml_reader::get_attribute(unsigned int i, string& name, string& value) {
    if (!xmlTextReaderMoveToAttributeNo(reader, i))
        return false;

    {
        auto xc = xmlTextReaderConstName(reader);

        if (!xc)
            name = "";
        else
            name = (char*)xc;
    }

    {
        auto xc = xmlTextReaderConstValue(reader);

        if (!xc)
            value = "";
        else
            value = (char*)xc;
    }

    return true;
}

bool xml_reader::is_empty() const {
    return xmlTextReaderIsEmptyElement(reader);
}
