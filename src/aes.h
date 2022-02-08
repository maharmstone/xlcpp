#pragma once

#include <stdint.h>
#include <stddef.h>

#define AES_BLOCKLEN 16 // Block length in bytes - AES is 128b block only
#define AES_KEYLEN 16   // Key length in bytes
#define AES_keyExpSize 176

struct AES_ctx {
    uint8_t RoundKey[AES_keyExpSize];
    uint8_t Iv[AES_BLOCKLEN];
};

void AES_init_ctx(struct AES_ctx* ctx, const uint8_t* key);
void AES_init_ctx_iv(struct AES_ctx* ctx, const uint8_t* key, const uint8_t* iv);
void AES_ECB_encrypt(const struct AES_ctx* ctx, uint8_t* buf);
void AES_ECB_decrypt(const struct AES_ctx* ctx, uint8_t* buf);
void AES_CBC_encrypt_buffer(struct AES_ctx* ctx, uint8_t* buf, size_t length);
void AES_CBC_decrypt_buffer(struct AES_ctx* ctx, uint8_t* buf, size_t length);
