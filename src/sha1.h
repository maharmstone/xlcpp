/*
SHA-1 in C
By Steve Reid <steve@edmweb.com>
100% Public Domain
*/

#pragma once

#include <array>
#include <span>

struct SHA1_CTX {
    constexpr SHA1_CTX();
    constexpr void update(std::span<const uint8_t> data);
    constexpr void finalize(std::array<uint8_t, 20>& digest);

    uint32_t state[5];
    uint32_t count[2];
    unsigned char buffer[64];
};

std::array<uint8_t, 20> sha1(std::span<const uint8_t> s);
