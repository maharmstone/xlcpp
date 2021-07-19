#include "xlcpp-pimpl.h"

using namespace std;

static bool __inline is_whitespace(char c) {
    return c == ' ' || c == '\t' || c == '\r' || c == '\n';
}

static void parse_attributes(const string_view& node, const function<bool(const string_view&, const xml_enc_string_view&)>& func) {
    auto s = node.substr(1, node.length() - 2);

    if (!s.empty() && s.back() == '/') {
        s.remove_suffix(1);
    }

    while (!s.empty() && !is_whitespace(s.front())) {
        s.remove_prefix(1);
    }

    while (!s.empty() && is_whitespace(s.front())) {
        s.remove_prefix(1);
    }

    while (!s.empty()) {
        auto av = s;

        auto eq = av.find_first_of('=');
        string_view n, v;

        if (eq == string::npos) {
            n = av;
            v = "";
        } else {
            n = av.substr(0, eq);
            v = av.substr(eq + 1);

            while (!v.empty() && is_whitespace(v.front())) {
                v = v.substr(1);
            }

            if (v.length() >= 2 && (v.front() == '"' || v.front() == '\'')) {
                auto c = v.front();

                v.remove_prefix(1);

                auto end = v.find_first_of(c);

                if (end != string::npos) {
                    v = v.substr(0, end);
                    av = av.substr(0, v.data() + v.length() - av.data() + 1);
                } else {
                    for (size_t i = 0; i < av.length(); i++) {
                        if (is_whitespace(av[i])) {
                            av = av.substr(0, i);
                            break;
                        }
                    }
                }
            }
        }

        while (!n.empty() && is_whitespace(n.back())) {
            n.remove_suffix(1);
        }

        if (!func(n, v))
            return;

        s.remove_prefix(av.length());

        while (!s.empty() && is_whitespace(s.front())) {
            s.remove_prefix(1);
        }
    }
}

bool xml_reader::read() {
    if (sv.empty())
        return false;

    // FIXME - CDATA
    // FIXME - comments

    if (type == xml_node::element && empty_tag)
        namespaces.pop_back();

    if (sv.front() != '<') { // text
        auto pos = sv.find_first_of('<');

        if (pos == string::npos) {
            node = sv;
            sv = "";
        } else {
            node = sv.substr(0, pos);
            sv = sv.substr(pos);
        }

        type = xml_node::whitespace;

        for (auto c : sv) {
            if (!is_whitespace(c)) {
                type = xml_node::text;
                break;
            }
        }
    } else {
        auto pos = sv.find_first_of('>');

        if (pos == string::npos) {
            node = sv;
            sv = "";
        } else {
            node = sv.substr(0, pos + 1);
            sv = sv.substr(pos + 1);
        }

        if (node.starts_with("<?xml"))
            type = xml_node::xml_declaration;
        else if (node.starts_with("</")) {
            type = xml_node::end_element;
            namespaces.pop_back();
        } else {
            type = xml_node::element;
            ns_list ns;

            parse_attributes(node, [&](const string_view& name, const xml_enc_string_view& value) {
                if (name.starts_with("xmlns:"))
                    ns.emplace_back(name.substr(6), value);
                else if (name == "xmlns")
                    ns.emplace_back("", value);

                return true;
            });

            namespaces.push_back(ns);

            empty_tag = node.ends_with("/>");
        }
    }

    return true;
}

enum xml_node xml_reader::node_type() const {
    return type;
}

bool xml_reader::is_empty() const {
    return type == xml_node::element && empty_tag;
}

void xml_reader::attributes_loop_raw(const function<bool(const string_view& local_name, const xml_enc_string_view& namespace_uri_raw,
                                                         const xml_enc_string_view& value_raw)>& func) const {
    if (type != xml_node::element)
        return;

    parse_attributes(node, [&](const string_view& name, const xml_enc_string_view& value_raw) {
        auto colon = name.find_first_of(':');

        if (colon == string::npos)
            return func(name, xml_enc_string_view{}, value_raw);

        auto prefix = name.substr(0, colon);

        for (auto it = namespaces.rbegin(); it != namespaces.rend(); it++) {
            for (const auto& v : *it) {
                if (v.first == prefix)
                    return func(name.substr(colon + 1), v.second, value_raw);
            }
        }

        return func(name.substr(colon + 1), xml_enc_string_view{}, value_raw);
    });
}

optional<xml_enc_string_view> xml_reader::get_attribute(const string_view& name, const string_view& ns) const {
    if (type != xml_node::element)
        return nullopt;

    optional<xml_enc_string_view> xesv;

    attributes_loop_raw([&](const string_view& local_name, const xml_enc_string_view& namespace_uri_raw,
                            const xml_enc_string_view& value_raw) {
        if (local_name == name && namespace_uri_raw.cmp(ns)) {
            xesv = value_raw;
            return false;
        }

        return true;
    });

    return xesv;
}

xml_enc_string_view xml_reader::namespace_uri_raw() const {
    auto tag = name();
    auto colon = tag.find_first_of(':');
    string_view prefix;

    if (colon != string::npos)
        prefix = tag.substr(0, colon);

    for (auto it = namespaces.rbegin(); it != namespaces.rend(); it++) {
        for (const auto& v : *it) {
            if (v.first == prefix)
                return v.second;
        }
    }

    return {};
}

string_view xml_reader::name() const {
    if (type != xml_node::element && type != xml_node::end_element)
        return "";

    auto tag = node.substr(type == xml_node::end_element ? 2 : 1);

    tag.remove_suffix(1);

    for (size_t i = 0; i < tag.length(); i++) {
        if (is_whitespace(tag[i])) {
            tag = tag.substr(0, i);
            break;
        }
    }

    return tag;
}

string_view xml_reader::local_name() const {
    if (type != xml_node::element && type != xml_node::end_element)
        return "";

    auto tag = name();
    auto pos = tag.find_first_of(':');

    if (pos == string::npos)
        return tag;
    else
        return tag.substr(pos + 1);
}

xml_enc_string_view xml_reader::value_raw() const {
    if (type != xml_node::text)
        return {};

    return node;
}

string xml_enc_string_view::decode() const {
    // FIXME - lt, gt, amp, quot, numeric, hex

    return string{sv};
}

bool xml_enc_string_view::cmp(const string_view& str) const {
    for (auto c : sv) {
        if (c == '&')
            return decode() == str;
    }

    return sv == str;
}
