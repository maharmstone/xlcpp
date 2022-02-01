#include <fmt/format.h>
#include <fstream>
#include "cfbf.h"
#include "utf16.h"
#include "sha1.h"

using namespace std;

static const uint64_t CFBF_SIGNATURE = 0xe11ab1a1e011cfd0;
static const uint32_t NOSTREAM = 0xffffffff;

struct structured_storage_header {
    uint64_t sig;
    uint8_t clsid[16];
    uint16_t minor_version;
    uint16_t major_version;
    uint16_t byte_order;
    uint16_t sector_shift;
    uint16_t mini_sector_shift;
    uint16_t reserved1;
    uint32_t reserved2;
    uint32_t num_sect_dir;
    uint32_t num_sect_fat;
    uint32_t sect_dir_start;
    uint32_t transaction_signature;
    uint32_t mini_sector_cutoff;
    uint32_t mini_fat_start;
    uint32_t num_sect_mini_fat;
    uint32_t sect_dif_start;
    uint32_t num_sect_dif;
    uint32_t sect_dif[109];
};

static_assert(sizeof(structured_storage_header) == 0x200);

enum class obj_type : uint8_t {
    STGTY_INVALID = 0,
    STGTY_STORAGE = 1,
    STGTY_STREAM = 2,
    STGTY_LOCKBYTES = 3,
    STGTY_PROPERTY = 4,
    STGTY_ROOT = 5
};

enum class tree_colour : uint8_t {
    red = 0,
    black = 1
};

#pragma pack(push,1)

struct dirent {
    char16_t name[32];
    uint16_t name_len;
    obj_type type;
    tree_colour colour;
    uint32_t sid_left_sibling;
    uint32_t sid_right_sibling;
    uint32_t sid_child;
    uint8_t clsid[16];
    uint32_t user_flags;
    uint64_t create_time;
    uint64_t modify_time;
    uint32_t sect_start;
    uint64_t size;
};

#pragma pack(pop)

static_assert(sizeof(dirent) == 0x80);

cfbf::cfbf(const filesystem::path& fn) {
    unique_handle hup{open(fn.string().c_str(), O_RDONLY)};

    m = make_unique<mmap>(hup.get());

    s = m->map();

    auto& ssh = *(structured_storage_header*)s.data();

    if (ssh.sig != CFBF_SIGNATURE)
        throw runtime_error("Incorrect signature.");

    fmt::print("sig = {:016x}\n", ssh.sig);
    fmt::print("clsid = {:x}\n", *(uint64_t*)ssh.clsid);
    fmt::print("minor_version = {:x}\n", ssh.minor_version);
    fmt::print("major_version = {:x}\n", ssh.major_version);
    fmt::print("byte_order = {:x}\n", ssh.byte_order);
    fmt::print("sector_shift = {:x}\n", ssh.sector_shift);
    fmt::print("mini_sector_shift = {:x}\n", ssh.mini_sector_shift);
    fmt::print("reserved1 = {:x}\n", ssh.reserved1);
    fmt::print("reserved2 = {:x}\n", ssh.reserved2);
    fmt::print("num_sect_dir = {:x}\n", ssh.num_sect_dir);
    fmt::print("num_sect_fat = {:x}\n", ssh.num_sect_fat);
    fmt::print("sect_dir_start = {:x}\n", ssh.sect_dir_start);
    fmt::print("transaction_signature = {:x}\n", ssh.transaction_signature);
    fmt::print("mini_sector_cutoff = {:x}\n", ssh.mini_sector_cutoff);
    fmt::print("mini_fat_start = {:x}\n", ssh.mini_fat_start);
    fmt::print("num_sect_mini_fat = {:x}\n", ssh.num_sect_mini_fat);
    fmt::print("sect_dif_start = {:x}\n", ssh.sect_dif_start);
    fmt::print("num_sect_dif = {:x}\n", ssh.num_sect_dif);

    for (unsigned int i = 0; i < ssh.num_sect_dif; i++) {
        fmt::print("sect_dif = {:x}\n", ssh.sect_dif[i]);
    }

    fmt::print("---\n");

    auto& de = *(dirent*)(s.data() + (ssh.sect_dir_start + 1) * (1 << ssh.sector_shift));

    if (de.type != obj_type::STGTY_ROOT)
        throw runtime_error("Root directory entry did not have type STGTY_ROOT.");

    add_entry("", 0);
}

void cfbf::add_entry(string_view path, uint32_t num) {
    auto& ssh = *(structured_storage_header*)s.data();
    auto dirents = s.subspan((ssh.sect_dir_start + 1) * (1 << ssh.sector_shift));
    auto& de = *(dirent*)(dirents.data() + (num * sizeof(dirent)));
    auto name = de.name_len >= sizeof(char16_t) && num != 0 ? utf16_to_utf8(u16string_view(de.name, (de.name_len / sizeof(char16_t)) - 1)) : "";

    entries.emplace_back(*this, de, string(path) + name);

    if (de.sid_child != NOSTREAM)
        add_entry(name + "/", de.sid_child);

    if (de.sid_right_sibling != NOSTREAM)
        add_entry(path, de.sid_right_sibling);
}

cfbf_entry::cfbf_entry(cfbf& file, dirent& de, string_view name) : file(file), de(de), name(name) {
}

uint32_t cfbf::next_sector(uint32_t sector) const {
    auto& ssh = *(structured_storage_header*)s.data();
    auto fat = (uint32_t*)(s.data() + ((ssh.sect_dif[0] + 1) << ssh.sector_shift));

    return fat[sector];
}

uint32_t cfbf::next_mini_sector(uint32_t sector) const {
    auto& ssh = *(structured_storage_header*)s.data();
    auto mini_fat = (uint32_t*)(s.data() + ((ssh.mini_fat_start + 1) << ssh.sector_shift));

    return mini_fat[sector];
}

size_t cfbf_entry::read(span<std::byte> buf, uint64_t off) const {
    auto& ssh = *(structured_storage_header*)file.s.data();

    if (off > de.size)
        return 0;

    if (off + buf.size() > de.size)
        buf = buf.subspan(0, de.size - off);

    size_t read = 0;
    auto sector = de.sect_start;

    if (de.size < ssh.mini_sector_cutoff) {
        auto sector_skip = off >> ssh.mini_sector_shift;

        for (unsigned int i = 0; i < sector_skip; i++) {
            sector = file.next_mini_sector(sector);
        }

        // FIXME - what if mini_stream more than 1 sector?
        auto mini_stream = file.s.subspan((file.entries[0].de.sect_start + 1) << ssh.sector_shift, 1 << ssh.sector_shift);

        do {
            auto src = mini_stream.subspan(sector << ssh.mini_sector_shift, 1 << ssh.mini_sector_shift);
            auto to_copy = min(src.size(), buf.size());

            memcpy(buf.data(), src.data(), to_copy);

            read += to_copy;
            buf = buf.subspan(to_copy);

            if (buf.empty())
                break;

            sector = file.next_mini_sector(sector);
        } while (true);
    } else {
        auto sector = de.sect_start;
        auto sector_skip = off >> ssh.sector_shift;

        for (unsigned int i = 0; i < sector_skip; i++) {
            sector = file.next_sector(sector);
        }

        do {
            auto src = file.s.subspan((sector + 1) << ssh.sector_shift, 1 << ssh.sector_shift);
            auto to_copy = min(src.size(), buf.size());

            memcpy(buf.data(), src.data(), to_copy);

            read += to_copy;
            buf = buf.subspan(to_copy);

            if (buf.empty())
                break;

            sector = file.next_sector(sector);
        } while (true);
    }

    return read;
}

static void cfbf_test(const filesystem::path& fn) {
    cfbf c(fn);

    for (unsigned int num = 0; const auto& e : c.entries) {
        fmt::print("{}\n", e.name);
        fmt::print("  type = {:x}\n", (unsigned int)e.de.type);
        fmt::print("  colour = {:x}\n", (unsigned int)e.de.colour);
        fmt::print("  sid_left_sibling = {:x}\n", e.de.sid_left_sibling);
        fmt::print("  sid_right_sibling = {:x}\n", e.de.sid_right_sibling);
        fmt::print("  sid_child = {:x}\n", e.de.sid_child);
        fmt::print("  clsid = {:x}\n", *(uint64_t*)e.de.clsid);
        fmt::print("  user_flags = {:x}\n", e.de.user_flags);
        fmt::print("  create_time = {:x}\n", e.de.create_time);
        fmt::print("  modify_time = {:x}\n", e.de.modify_time);
        fmt::print("  sect_start = {:x}\n", e.de.sect_start);
        fmt::print("  size = {:x}\n", e.de.size);
        fmt::print("--\n");

        if (num == 0) { // root
            num++;
            continue;
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
}

int main() {
    auto hash = sha1(span<uint8_t>());

    for (auto c : hash) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");

    try {
        cfbf_test("../password.xlsx");
    } catch (const exception& e) {
        fmt::print(stderr, "{}\n", e.what());
        return 1;
    }

    return 0;
}
