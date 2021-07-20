#include "xlcpp-pimpl.h"

using namespace std;

string xml_writer::dump() const {
    return buf;
}

void xml_writer::start_document() {
    buf += "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
    empty_tag = false;
}

static string xml_escape(const string_view& s, bool att) {
    string ret;

    ret.reserve(s.length());

    for (auto c : s) {
        if (c == '<')
            ret += "&lt;";
        else if (c == '>')
            ret += "&gt;";
        else if (c == '&')
            ret += "&amp;";
        else if (c == '"' && att)
            ret += "&quot;";
        else
            ret += c;
    }

    return ret;
}

void xml_writer::start_element(const string& tag, const unordered_map<string, string>& namespaces) {
    if (empty_tag)
        buf += ">";

    buf += "<" + tag;
    tags.push(tag);

    empty_tag = true;

    for (const auto& ns : namespaces) {
        buf += " xmlns";

        if (!ns.first.empty())
            buf += ":" + ns.first;

        buf += "=\"";
        buf += xml_escape(ns.second, true);
        buf += "\"";
    }
}

void xml_writer::end_element() {
    if (empty_tag) {
        buf += "/>";
        empty_tag = false;
    } else
        buf += "</" + tags.top() + ">";

    tags.pop();
}

void xml_writer::text(const string& s) {
    if (empty_tag) {
        buf += ">";
        empty_tag = false;
    }

    buf += xml_escape(s, false);
}

void xml_writer::attribute(const string& name, const string& value) {
    buf += " " + name + "=\"" + xml_escape(value, true) + "\"";
}
