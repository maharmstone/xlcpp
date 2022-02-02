#include <fmt/format.h>
#include <fstream>
#include "cfbf.h"
#include "utf16.h"
#include "sha1.h"
#include "aes.h"

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

const dirent& cfbf::find_dirent(uint32_t num) {
    auto& ssh = *(structured_storage_header*)s.data();

    auto dirents_per_sector = (1 << ssh.sector_shift) / sizeof(dirent);
    auto sector_skip = num / dirents_per_sector;
    auto sector = ssh.sect_dir_start;

    while (sector_skip > 0) {
        sector = next_sector(sector);
        sector_skip--;
    }

    return *(dirent*)(s.data() + ((sector + 1) << ssh.sector_shift) + ((num % dirents_per_sector) * sizeof(dirent)));
}

void cfbf::add_entry(string_view path, uint32_t num) {
    const auto& de = find_dirent(num);

    auto name = de.name_len >= sizeof(char16_t) && num != 0 ? utf16_to_utf8(u16string_view(de.name, (de.name_len / sizeof(char16_t)) - 1)) : "";

    entries.emplace_back(*this, de, string(path) + name);

    if (de.sid_child != NOSTREAM)
        add_entry(name + "/", de.sid_child);

    if (de.sid_right_sibling != NOSTREAM)
        add_entry(path, de.sid_right_sibling);
}

cfbf_entry::cfbf_entry(cfbf& file, const dirent& de, string_view name) : file(file), de(de), name(name) {
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

static array<uint8_t, 16> generate_key(u16string_view password, span<const uint8_t> salt) {
    array<uint8_t, 20> h;

    {
        SHA1_CTX ctx;

        ctx.update(salt);
        ctx.update(span((uint8_t*)password.data(), password.size() * sizeof(char16_t)));

        ctx.finalize(h);
    }

    for (uint32_t i = 0; i < 50000; i++) {
        SHA1_CTX ctx;

        ctx.update(span((uint8_t*)&i, sizeof(uint32_t)));
        ctx.update(h);

        ctx.finalize(h);
    }

    {
        SHA1_CTX ctx;
        uint32_t block = 0;

        ctx.update(h);
        ctx.update(span((uint8_t*)&block, sizeof(uint32_t)));

        ctx.finalize(h);
    }

    array<uint8_t, 64> buf1 = {
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36
    };

    for (unsigned int i = 0; auto c : h) {
        buf1[i] ^= c;
        i++;
    }

    auto x1 = sha1(buf1);

    array<uint8_t, 16> ret;
    memcpy(ret.data(), x1.data(), ret.size());

    return ret;
}

#pragma pack(push, 1)

struct encryption_info {
    uint16_t major;
    uint16_t minor;
    uint32_t flags;
    uint32_t header_size;
};

struct encryption_header {
    uint32_t flags;
    uint32_t size_extra;
    uint32_t alg_id;
    uint32_t alg_id_hash;
    uint32_t key_size;
    uint32_t provider_type;
    uint32_t reserved1;
    uint32_t reserved2;
    char16_t csp_name[0];
};

#pragma pack(pop)

static const uint32_t ALG_ID_AES_128 = 0x660e;
static const uint32_t ALG_ID_SHA_1 = 0x8004;

void cfbf::check_password(u16string_view password, span<const uint8_t> salt,
                          span<const uint8_t> encrypted_verifier,
                          span<const uint8_t> encrypted_verifier_hash) {
    auto key = generate_key(password, salt);
    AES_ctx ctx;
    array<uint8_t, 16> verifier;
    array<uint8_t, 32> verifier_hash;

    if (encrypted_verifier.size() != verifier.size())
        throw formatted_error("encrypted_verifier.size() was {}, expected {}", encrypted_verifier.size(), verifier.size());

    if (encrypted_verifier_hash.size() != verifier_hash.size())
        throw formatted_error("encrypted_verifier_hash.size() was {}, expected {}", encrypted_verifier_hash.size(), verifier_hash.size());

    AES_init_ctx(&ctx, key.data());

    memcpy(verifier.data(), encrypted_verifier.data(), encrypted_verifier.size());

    AES_ECB_decrypt(&ctx, verifier.data());

    fmt::print("verifier = ");
    for (auto c : verifier) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");

    memcpy(verifier_hash.data(), encrypted_verifier_hash.data(), encrypted_verifier_hash.size());

    AES_ECB_decrypt(&ctx, verifier_hash.data());
    AES_ECB_decrypt(&ctx, verifier_hash.data() + 16);

    fmt::print("verifier hash = ");
    for (auto c : verifier_hash) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");

    auto hash = sha1(verifier);

    if (memcmp(hash.data(), verifier_hash.data(), hash.size()))
        throw runtime_error("Incorrect password.");

    this->key = key;
}

void cfbf::parse_enc_info(span<const uint8_t> enc_info, u16string_view password) {
    if (enc_info.size() < sizeof(encryption_info))
        throw formatted_error("EncryptionInfo was {} bytes, expected at least {}", enc_info.size(), sizeof(encryption_info));

    auto& ei = *(encryption_info*)enc_info.data();

    if (ei.major != 3 || ei.minor != 2)
        throw formatted_error("Unsupported EncryptionInfo version {}.{}", ei.major, ei.minor);

    if (ei.flags != 0x24) // AES
        throw formatted_error("Unsupported EncryptionInfo flags {:x}", ei.flags);

    fmt::print("major = {:x}\n", ei.major);
    fmt::print("minor = {:x}\n", ei.minor);
    fmt::print("flags = {:x}\n", ei.flags);
    fmt::print("header_size = {:x}\n", ei.header_size);

    if (ei.header_size < offsetof(encryption_header, csp_name))
        throw formatted_error("Encryption header was {} bytes, expected at least {}", ei.header_size, offsetof(encryption_header, csp_name));

    if (ei.header_size > enc_info.size() - sizeof(encryption_info))
        throw formatted_error("Encryption header was {} bytes, but only {} remaining", ei.header_size, enc_info.size() - sizeof(encryption_info));

    auto& h = *(encryption_header*)(enc_info.data() + sizeof(encryption_info));

    fmt::print("flags = {:x}\n", h.flags);
    fmt::print("size_extra = {:x}\n", h.size_extra);
    fmt::print("alg_id = {:x}\n", h.alg_id);
    fmt::print("alg_id_hash = {:x}\n", h.alg_id_hash);
    fmt::print("key_size = {:x}\n", h.key_size);
    fmt::print("provider_type = {:x}\n", h.provider_type);

    if (h.alg_id != ALG_ID_AES_128)
        throw formatted_error("Unsupported algorithm ID {:x}", h.alg_id);

    if (h.alg_id_hash != ALG_ID_SHA_1 && h.alg_id_hash != 0)
        throw formatted_error("Unsupported hash algorithm ID {:x}", h.alg_id_hash);

    if (h.key_size != 128)
        throw formatted_error("Key size was {}, expected 128", h.key_size);

    auto sp = enc_info.subspan(sizeof(encryption_info) + ei.header_size);

    if (sp.size() < sizeof(uint32_t))
        throw runtime_error("Malformed EncryptionInfo");

    auto salt_size = *(uint32_t*)sp.data();
    sp = sp.subspan(sizeof(uint32_t));

    if (sp.size() < salt_size)
        throw runtime_error("Malformed EncryptionInfo");

    auto salt = sp.subspan(0, salt_size);
    sp = sp.subspan(salt_size);

    if (sp.size() < 16)
        throw runtime_error("Malformed EncryptionInfo");

    auto encrypted_verifier = sp.subspan(0, 16);
    sp = sp.subspan(16);

    if (sp.size() < sizeof(uint32_t))
        throw runtime_error("Malformed EncryptionInfo");

    // skip verifier_hash_size
    sp = sp.subspan(sizeof(uint32_t));

    if (sp.size() < 32)
        throw runtime_error("Malformed EncryptionInfo");

    auto encrypted_verifier_hash = sp.subspan(0, 32);

    check_password(password, salt, encrypted_verifier, encrypted_verifier_hash); // throws if wrong

    memcpy(this->salt.data(), salt.data(), this->salt.size());
}

void cfbf::decrypt(span<uint8_t> enc_package) {
    if (enc_package.size() < sizeof(uint64_t))
        throw formatted_error("EncryptedPackage was {} bytes, expected at least {}", enc_package.size(), sizeof(uint64_t));

    auto size = *(uint64_t*)enc_package.data();

    enc_package = enc_package.subspan(sizeof(uint64_t));

    if (enc_package.size() < size)
        throw formatted_error("EncryptedPackage was {} bytes, expected at least {}", enc_package.size() + sizeof(uint64_t), size + sizeof(uint64_t));

    AES_ctx ctx;
    auto buf = enc_package;

    AES_init_ctx(&ctx, key.data());

    while (!buf.empty()) {
        AES_ECB_decrypt(&ctx, buf.data());

        buf = buf.subspan(16);
    }

    ofstream f("plaintext");

    f.write((char*)enc_package.data(), enc_package.size()); // FIXME - throw exception if fails
}

static void cfbf_test(const filesystem::path& fn) {
    cfbf c(fn);
    string enc_info, enc_package;

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

        if (e.name == "/EncryptionInfo") {
            enc_info.resize(e.de.size);

            uint64_t off = 0;
            auto buf = span((std::byte*)enc_info.data(), enc_info.size());

            while (true) {
                auto size = e.read(buf, off);

                if (size == 0)
                    break;

                off += size;
            }
        } else if (e.name == "/EncryptedPackage") {
            enc_package.resize(e.de.size);

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

static void test_key() {
    const u16string_view password = u"Password1234_";
    array<uint8_t, 16> salt{0xe8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8, 0x30, 0xef};

    auto key = generate_key(password, salt);

    for (auto c : key) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");
}

int main() {
    test_sha1();

    test_key();

    try {
        cfbf_test("../password.xlsx");
    } catch (const exception& e) {
        fmt::print(stderr, "{}\n", e.what());
        return 1;
    }

    return 0;
}
