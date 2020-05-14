cmake_minimum_required(VERSION 3.10)

include(GNUInstallDirs)

project(xlcpp)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(LibXml2)

add_definitions(-DXLCPP_EXPORT)

set(SRC_FILES xlcpp.cpp)

add_library(xlcpp SHARED ${SRC_FILES})
add_library(xlcppstatic STATIC ${SRC_FILES})
target_link_libraries(xlcpp archive)

target_compile_options(xlcpp PRIVATE
	$<$<OR:$<CXX_COMPILER_ID:Clang>,$<CXX_COMPILER_ID:AppleClang>,$<CXX_COMPILER_ID:GNU>>:
		-Wall>
	$<$<CXX_COMPILER_ID:MSVC>:
		/W4>)

target_include_directories(xlcpp PUBLIC "${LIBXML2_INCLUDE_DIR}")
target_include_directories(xlcppstatic PUBLIC "${LIBXML2_INCLUDE_DIR}")

install(TARGETS xlcpp DESTINATION "${CMAKE_INSTALL_FULL_LIBDIR}")
install(TARGETS xlcppstatic DESTINATION "${CMAKE_INSTALL_FULL_LIBDIR}")
install(FILES xlcpp.h DESTINATION "${CMAKE_INSTALL_FULL_INCLUDEDIR}")

add_executable(xltest xltest.cpp)
target_include_directories(xltest PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}")
target_link_libraries(xltest xlcpp)

