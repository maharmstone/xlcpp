#pragma once

#include <string>
#include <string_view>
#include <span>
#include <unistd.h>
#include <fcntl.h>

#ifdef _WIN32
using fd_t = HANDLE;
#else
using fd_t = int;

class unique_handle {
public:
    unique_handle() : fd(0) {
    }

    explicit unique_handle(int fd) : fd(fd) {
    }

    unique_handle(unique_handle&& that) noexcept {
        fd = that.fd;
        that.fd = 0;
    }

    unique_handle(const unique_handle&) = delete;
    unique_handle& operator=(const unique_handle&) = delete;

    unique_handle& operator=(unique_handle&& that) noexcept {
        if (fd > 0)
            close(fd);

        fd = that.fd;
        that.fd = 0;

        return *this;
    }

    ~unique_handle() {
        if (fd <= 0)
            return;

        close(fd);
    }

    void reset(int new_fd) noexcept {
        if (fd > 0)
            close(fd);

        fd = new_fd;
    }

    int get() const noexcept {
        return fd;
    }

private:
    int fd;
};
#endif

class errno_error : public std::exception {
public:
    errno_error(const std::string_view& function, int en);

    const char* what() const noexcept {
        return msg.c_str();
    }

private:
    std::string msg;
};

class mmap {
public:
	explicit mmap(fd_t h);
	~mmap();
	std::span<const std::byte> map();

	size_t filesize;

private:
#ifdef _WIN32
	HANDLE h = INVALID_HANDLE_VALUE;
	HANDLE mh = INVALID_HANDLE_VALUE;
#endif
	void* ptr = nullptr;
};
