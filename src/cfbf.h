#pragma once

#include "mmap.h"
#include <string>
#include <string_view>
#include <filesystem>
#include <vector>
#include <span>

class cfbf;
struct dirent;

class cfbf_entry {
public:
    cfbf_entry(cfbf& file, dirent& de, std::string_view name);
    size_t read(std::span<std::byte> buf, uint64_t off) const;

    cfbf& file;
    dirent& de;
    std::string name;
};

class cfbf {
public:
    cfbf(const std::filesystem::path& fn);
    uint32_t next_sector(uint32_t sector) const;

    std::vector<cfbf_entry> entries;
    std::span<const std::byte> s;

private:
    void add_entry(std::string_view path, uint32_t num);

    std::unique_ptr<mmap> m;
};
