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

    cfbf& file;
    dirent& de;
    std::string name;
};

class cfbf {
public:
    cfbf(const std::filesystem::path& fn);

    std::vector<cfbf_entry> entries;

private:
    void add_entry(std::string_view path, uint32_t num);

    std::unique_ptr<mmap> m;
    std::span<const std::byte> s;
};
