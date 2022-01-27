#include <fmt/format.h>
#include <filesystem>
#include "mmap.h"
#include "utf16.h"

using namespace std;

static const uint64_t CFBF_SIGNATURE = 0xe11ab1a1e011cfd0;

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

static void cfbf_test(const filesystem::path& fn) {
    unique_handle hup{open(fn.string().c_str(), O_RDONLY)};

    mmap m(hup.get());

    auto s = m.map();

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

    fmt::print("name = \"{}\"\n", de.name_len >= sizeof(char16_t) ? utf16_to_utf8(u16string_view(de.name, (de.name_len / sizeof(char16_t)) - 1)) : "");
    fmt::print("type = {:x}\n", (unsigned int)de.type);
    fmt::print("colour = {:x}\n", (unsigned int)de.colour);
    fmt::print("sid_left_sibling = {:x}\n", de.sid_left_sibling);
    fmt::print("sid_right_sibling = {:x}\n", de.sid_right_sibling);
    fmt::print("sid_child = {:x}\n", de.sid_child);
    fmt::print("clsid = {:x}\n", *(uint64_t*)de.clsid);
    fmt::print("user_flags = {:x}\n", de.user_flags);
    fmt::print("create_time = {:x}\n", de.create_time);
    fmt::print("modify_time = {:x}\n", de.modify_time);
    fmt::print("sect_start = {:x}\n", de.sect_start);
    fmt::print("size = {:x}\n", de.size);
}

int main() {
    try {
        cfbf_test("../password.xlsx");
    } catch (const exception& e) {
        fmt::print(stderr, "{}\n", e.what());
        return 1;
    }

    return 0;
}
