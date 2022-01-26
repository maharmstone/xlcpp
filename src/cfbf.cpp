#include <fmt/format.h>
#include <fstream>

using namespace std;

static const uint64_t CFBF_SIGNATURE = 0xe11ab1a1e011cfd0;

struct structured_storage_header {
    uint64_t sig;
    uint8_t clsid[16];
    uint16_t minor_version;
    uint16_t dll_version;
    uint16_t byte_order;
    uint16_t sector_shift;
    uint16_t mini_sector_shift;
    uint16_t reserved1;
    uint32_t reserved2;
    uint32_t sect_dir;
    uint32_t num_sect_fat;
    uint32_t num_sect_dir_start;
    uint32_t transaction_signature;
    uint32_t mini_sector_cutoff;
    uint32_t mini_fat_start;
    uint32_t num_sect_mini_fat;
    uint32_t sect_dif_start;
    uint32_t num_sect_dif;
    uint32_t sect_fat[109];
};

static_assert(sizeof(structured_storage_header) == 512);

static void cfbf_test(const string& fn) {
    structured_storage_header ssh;

    ifstream f(fn);

    if (!f.good())
        throw runtime_error("Failed to open file.");

    f.read((char*)&ssh, sizeof(ssh));

    if (ssh.sig != CFBF_SIGNATURE)
        throw runtime_error("Incorrect signature.");

    fmt::print("sig = {:016x}\n", ssh.sig);
    fmt::print("clsid = {:x}\n", *(uint64_t*)ssh.clsid);
    fmt::print("minor_version = {:x}\n", ssh.minor_version);
    fmt::print("dll_version = {:x}\n", ssh.dll_version);
    fmt::print("byte_order = {:x}\n", ssh.byte_order);
    fmt::print("sector_shift = {:x}\n", ssh.sector_shift);
    fmt::print("mini_sector_shift = {:x}\n", ssh.mini_sector_shift);
    fmt::print("reserved1 = {:x}\n", ssh.reserved1);
    fmt::print("reserved2 = {:x}\n", ssh.reserved2);
    fmt::print("sect_dir = {:x}\n", ssh.sect_dir);
    fmt::print("num_sect_fat = {:x}\n", ssh.num_sect_fat);
    fmt::print("num_sect_dir_start = {:x}\n", ssh.num_sect_dir_start);
    fmt::print("transaction_signature = {:x}\n", ssh.transaction_signature);
    fmt::print("mini_sector_cutoff = {:x}\n", ssh.mini_sector_cutoff);
    fmt::print("mini_fat_start = {:x}\n", ssh.mini_fat_start);
    fmt::print("num_sect_mini_fat = {:x}\n", ssh.num_sect_mini_fat);
    fmt::print("sect_dif_start = {:x}\n", ssh.sect_dif_start);
    fmt::print("num_sect_dif = {:x}\n", ssh.num_sect_dif);

    for (unsigned int i = 0; i < ssh.num_sect_fat; i++) {
        fmt::print("sect_fat = {:x}\n", ssh.sect_fat[i]);
    }
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
