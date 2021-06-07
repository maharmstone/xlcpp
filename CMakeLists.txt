cmake_minimum_required(VERSION 3.10)

include(GNUInstallDirs)
include(CMakePackageConfigHelpers)

project(xlcpp)

option(BUILD_STATIC "Build static library" ON)
option(BUILD_SAMPLE "Build sample program" ON)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)

find_package(LibXml2 REQUIRED)
find_package(LibArchive REQUIRED)
find_package(fmt REQUIRED)

add_definitions(-DXLCPP_EXPORT)

set(SRC_FILES src/xlcpp.cpp
    src/xml-writer.cpp
    src/xml-reader.cpp)

add_library(xlcpp SHARED ${SRC_FILES})
target_link_libraries(xlcpp ${LIBXML2_LIBRARIES})
target_link_libraries(xlcpp ${LibArchive_LIBRARIES})
target_link_libraries(xlcpp fmt::fmt)

target_compile_options(xlcpp PRIVATE
    $<$<OR:$<CXX_COMPILER_ID:Clang>,$<CXX_COMPILER_ID:AppleClang>,$<CXX_COMPILER_ID:GNU>>:
        -Wall -Wno-trigraphs>
    $<$<CXX_COMPILER_ID:MSVC>:
        /W4>)

if(WIN32)
    add_definitions(-D_WIN32_WINNT=0x0600 -DNOMINMAX)
endif()

target_include_directories(xlcpp PUBLIC "${LibArchive_INCLUDE_DIR}")
target_include_directories(xlcpp PUBLIC "${LIBXML2_INCLUDE_DIR}")

set_target_properties(xlcpp PROPERTIES PUBLIC_HEADER src/xlcpp.h)

install(TARGETS xlcpp
    EXPORT xlcpp-targets
    RUNTIME DESTINATION "${CMAKE_INSTALL_BINDIR}"
    ARCHIVE DESTINATION "${CMAKE_INSTALL_LIBDIR}"
    LIBRARY DESTINATION "${CMAKE_INSTALL_LIBDIR}"
    PUBLIC_HEADER DESTINATION "${CMAKE_INSTALL_INCLUDEDIR}"
    PRIVATE_HEADER DESTINATION "${CMAKE_INSTALL_INCLUDEDIR}"
)

if(BUILD_STATIC)
    add_library(xlcppstatic STATIC ${SRC_FILES})
    target_include_directories(xlcppstatic PUBLIC "${LibArchive_INCLUDE_DIR}")
    target_include_directories(xlcppstatic PUBLIC "${LIBXML2_INCLUDE_DIR}")
    install(TARGETS xlcppstatic DESTINATION "${CMAKE_INSTALL_FULL_LIBDIR}")
endif()

if(BUILD_SAMPLE)
    add_executable(xltest src/xltest.cpp)
    target_include_directories(xltest PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}/src")
    target_link_libraries(xltest xlcpp)
    install(TARGETS xltest DESTINATION "${CMAKE_INSTALL_FULL_BINDIR}")
endif()

install(EXPORT xlcpp-targets DESTINATION lib/cmake/xlcpp)

configure_package_config_file(
    "xlcppConfig.cmake.in" "${CMAKE_CURRENT_BINARY_DIR}/xlcppConfig.cmake"
    INSTALL_DESTINATION "lib/cmake/xlcpp"
)

install(FILES "${CMAKE_CURRENT_BINARY_DIR}/xlcppConfig.cmake" DESTINATION "lib/cmake/xlcpp")
install(FILES "${CMAKE_CURRENT_BINARY_DIR}/libxlcpp.pdb" DESTINATION "${CMAKE_INSTALL_BINDIR}" OPTIONAL)
