#include <fstream>
#include "cfbf.h"
#include "sha1.h"

using namespace std;

static void cfbf_test(const filesystem::path& fn) {
#ifdef _WIN32
    unique_handle hup{CreateFileW((LPCWSTR)fn.u16string().c_str(), FILE_READ_DATA | DELETE, FILE_SHARE_READ, nullptr, OPEN_EXISTING,
                                  FILE_ATTRIBUTE_NORMAL, nullptr)};
    if (hup.get() == INVALID_HANDLE_VALUE)
        throw last_error("CreateFile", GetLastError());
#else
    unique_handle hup{open(fn.string().c_str(), O_RDONLY)};

    if (hup.get() == -1)
        throw formatted_error("open failed (errno = {})", errno);
#endif

    mmap m(hup.get());

    auto mem = m.map();

    cfbf c(mem);
    string enc_info, enc_package;

    for (unsigned int num = 0; const auto& e : c.entries) {
//         fmt::print("{}\n", e.name);
//         fmt::print("  type = {:x}\n", (unsigned int)e.de.type);
//         fmt::print("  colour = {:x}\n", (unsigned int)e.de.colour);
//         fmt::print("  sid_left_sibling = {:x}\n", e.de.sid_left_sibling);
//         fmt::print("  sid_right_sibling = {:x}\n", e.de.sid_right_sibling);
//         fmt::print("  sid_child = {:x}\n", e.de.sid_child);
//         fmt::print("  clsid = {:x}\n", *(uint64_t*)e.de.clsid);
//         fmt::print("  user_flags = {:x}\n", e.de.user_flags);
//         fmt::print("  create_time = {:x}\n", e.de.create_time);
//         fmt::print("  modify_time = {:x}\n", e.de.modify_time);
//         fmt::print("  sect_start = {:x}\n", e.de.sect_start);
//         fmt::print("  size = {:x}\n", e.de.size);
//         fmt::print("--\n");

        if (num == 0) { // root
            num++;
            continue;
        }

        if (e.name == "/EncryptionInfo") {
            enc_info.resize(e.get_size());

            uint64_t off = 0;
            auto buf = span((std::byte*)enc_info.data(), enc_info.size());

            while (true) {
                auto size = e.read(buf, off);

                if (size == 0)
                    break;

                off += size;
            }
        } else if (e.name == "/EncryptedPackage") {
            enc_package.resize(e.get_size());

            uint64_t off = 0;
            auto buf = span((std::byte*)enc_package.data(), enc_package.size());

            while (true) {
                auto size = e.read(buf, off);

                if (size == 0)
                    break;

                off += size;
            }
        }

        ofstream f("file" + to_string(num));

        uint64_t off = 0;
        std::byte buf[4096];

        while (true) {
            auto size = e.read(buf, off);

            if (size == 0)
                break;

            f.write((char*)buf, size); // FIXME - throw exception if fails

            off += size;
        }

        num++;
    }

    if (enc_info.empty())
        throw runtime_error("EncryptionInfo not found.");

    c.parse_enc_info(span((uint8_t*)enc_info.data(), enc_info.size()), u"password");
    c.decrypt(span((uint8_t*)enc_package.data(), enc_package.size()));
}

static void test_sha1() {
    auto hash = sha1(span<uint8_t>());

    for (auto c : hash) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");

    auto sv = string_view("abc");

    {
        SHA1_CTX ctx;

        for (auto c : sv) {
            ctx.update(span((uint8_t*)&c, 1));
        }

        ctx.finalize(hash);
    }

    for (auto c : hash) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");
}

#if 0
static void test_key() {
    const u16string_view password = u"Password1234_";
    array<uint8_t, 16> salt{0xe8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8, 0x30, 0xef};

    auto key = generate_key(password, salt);

    for (auto c : key) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");
}
#endif

int main() {
    test_sha1();

//     test_key();

    try {
        cfbf_test("../password.xlsx");
    } catch (const exception& e) {
        fmt::print(stderr, "{}\n", e.what());
        return 1;
    }

    return 0;
}
