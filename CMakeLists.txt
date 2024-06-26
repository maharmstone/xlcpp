cmake_minimum_required(VERSION 3.17)

project(xlcpp
    LANGUAGES CXX
)

include(GNUInstallDirs)
include(CMakePackageConfigHelpers)

option(BUILD_STATIC "Build static library" ON)
option(BUILD_SAMPLE "Build sample program" ON)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)

find_package(LibArchive REQUIRED)

add_definitions(-DXLCPP_EXPORT)

set(SRC_FILES src/xlcpp.cpp
    src/xml-writer.cpp
    src/xml-reader.cpp
    src/mmap.cpp
    src/cfbf.cpp
    src/aes.cpp
    src/sha1.cpp
    src/b64.cpp
    src/sha512.cpp
    src/xlsb.cpp)

add_library(xlcpp SHARED ${SRC_FILES})
target_link_libraries(xlcpp LibArchive::LibArchive)

target_compile_options(xlcpp PRIVATE
    $<$<OR:$<CXX_COMPILER_ID:Clang>,$<CXX_COMPILER_ID:AppleClang>,$<CXX_COMPILER_ID:GNU>>:
        -Wall -Wno-trigraphs -Wconversion>
    $<$<CXX_COMPILER_ID:MSVC>:
        /W4>)

if(WIN32)
    add_definitions(-D_WIN32_WINNT=0x0600 -DNOMINMAX)
endif()

set_target_properties(xlcpp PROPERTIES PUBLIC_HEADER src/xlcpp.h)

target_include_directories(xlcpp PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/src>
    $<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>
)

install(TARGETS xlcpp
    EXPORT xlcpp-targets
    RUNTIME DESTINATION "${CMAKE_INSTALL_BINDIR}"
    ARCHIVE DESTINATION "${CMAKE_INSTALL_LIBDIR}"
    LIBRARY DESTINATION "${CMAKE_INSTALL_LIBDIR}"
    PUBLIC_HEADER DESTINATION "${CMAKE_INSTALL_INCLUDEDIR}"
    PRIVATE_HEADER DESTINATION "${CMAKE_INSTALL_INCLUDEDIR}"
)

if(MSVC)
    install(FILES $<TARGET_PDB_FILE:xlcpp> DESTINATION bin OPTIONAL)
endif()

if(BUILD_STATIC)
    add_library(xlcppstatic STATIC ${SRC_FILES})
    target_link_libraries(xlcppstatic LibArchive::LibArchive)
    install(TARGETS xlcppstatic DESTINATION "${CMAKE_INSTALL_FULL_LIBDIR}")
endif()

if(BUILD_SAMPLE)
    add_executable(xltest src/xltest.cpp)
    target_include_directories(xltest PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}/src")
    target_link_libraries(xltest xlcpp)
    install(TARGETS xltest DESTINATION "${CMAKE_INSTALL_FULL_BINDIR}")

    if(MSVC)
        install(FILES $<TARGET_PDB_FILE:xltest> DESTINATION bin OPTIONAL)
    endif()

    # TESTING
    add_executable(cfbf src/cfbf.cpp src/mmap.cpp src/sha1.cpp src/aes.cpp src/cfbf-test.cpp src/xml-reader.cpp src/b64.cpp src/sha512.cpp)
    target_include_directories(cfbf PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}/src")
    target_link_libraries(cfbf xlcpp)
    install(TARGETS cfbf DESTINATION "${CMAKE_INSTALL_FULL_BINDIR}")

    if(MSVC)
        install(FILES $<TARGET_PDB_FILE:cfbf> DESTINATION bin OPTIONAL)
    endif()

    target_compile_options(cfbf PRIVATE
        $<$<OR:$<CXX_COMPILER_ID:Clang>,$<CXX_COMPILER_ID:AppleClang>,$<CXX_COMPILER_ID:GNU>>:
            -Wall -Wno-trigraphs -Wconversion>
        $<$<CXX_COMPILER_ID:MSVC>:
            /W4>)
endif()

install(EXPORT xlcpp-targets DESTINATION lib/cmake/xlcpp)

configure_package_config_file(
    "xlcppConfig.cmake.in" "${CMAKE_CURRENT_BINARY_DIR}/xlcppConfig.cmake"
    INSTALL_DESTINATION "lib/cmake/xlcpp"
)

install(FILES "${CMAKE_CURRENT_BINARY_DIR}/xlcppConfig.cmake" DESTINATION "lib/cmake/xlcpp")
install(FILES "${CMAKE_CURRENT_BINARY_DIR}/libxlcpp.pdb" DESTINATION "${CMAKE_INSTALL_BINDIR}" OPTIONAL)

set(CPACK_PACKAGE_NAME xlcpp-dev)
set(CPACK_DEBIAN_PACKAGE_MAINTAINER "Mark Harmstone")
set(CPACK_DEBIAN_PACKAGE_DEPENDS "libarchive-dev")
include(CPack)
