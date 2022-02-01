#pragma once

#include "mmap.h"
#include <string>
#include <string_view>
#include <filesystem>
#include <vector>
#include <span>
#include <fmt/format.h>
#include <fmt/compile.h>

class _formatted_error : public std::exception {
public:
    template<typename T, typename... Args>
    _formatted_error(T&& s, Args&&... args) {
        msg = fmt::format(s, std::forward<Args>(args)...);
    }

    const char* what() const noexcept {
        return msg.c_str();
    }

private:
    std::string msg;
};

#define formatted_error(s, ...) _formatted_error(FMT_COMPILE(s), ##__VA_ARGS__)

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
    uint32_t next_mini_sector(uint32_t sector) const;

    std::vector<cfbf_entry> entries;
    std::span<const std::byte> s;

private:
    void add_entry(std::string_view path, uint32_t num);

    std::unique_ptr<mmap> m;
};
