cmake_minimum_required(VERSION 3.16)

project(XLSX LANGUAGES CXX)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_WINDOWS_EXPORT_ALL_SYMBOLS ON)

if(CMAKE_VERSION VERSION_LESS "3.7.0")
    set(CMAKE_INCLUDE_CURRENT_DIR ON)
endif()

if(WIN32 AND CMAKE_CXX_COMPILER_ID MATCHES "MSVC")
  add_definitions("/wd4251")
  add_definitions("/wd4244")
  add_definitions("/wd4267")
  add_definitions("/wd4996")
  add_definitions("/wd4723")
  add_definitions("/wd4101")
  add_definitions("/wd4018")
  add_definitions("/wd4068")
  add_definitions("/wd4834")
  add_definitions("/wd4334")
  add_definitions("/wd4324")
  add_definitions("/wd4242")
  add_definitions("/wd4702")
  add_definitions("/wd4245")
else()
    add_compile_options(-O3)
endif()

#==============================================================================================================
# include CMake macro
#==============================================================================================================
include("../cmake/collect_source_code_and_resources.cmake")

#==============================================================================================================
# OpenXLSX
#==============================================================================================================
set(OPENXLSX_LIBRARY_TYPE "SHARED" CACHE STRING "Set the library type to SHARED or STATIC")
set(OPENXLSX_ENABLE_LTO OFF CACHE BOOL "Enables LTO for OpenXLSX" FORCE)
set(OPENXLSX_CREATE_DOCS OFF CACHE BOOL "Enables LTO for OpenXLSX" FORCE)
set(OPENXLSX_BUILD_SAMPLES OFF CACHE BOOL "Enables LTO for OpenXLSX" FORCE)
set(OPENXLSX_BUILD_TESTS OFF CACHE BOOL "Enables LTO for OpenXLSX" FORCE)
set(OPENXLSX_BUILD_BENCHMARKS OFF CACHE BOOL "Enables LTO for OpenXLSX" FORCE)
set(OPENXLSX_ENABLE_LIBZIP OFF CACHE BOOL "Enables LTO for OpenXLSX" FORCE)
add_subdirectory("kernel/tools/OpenXLSX/")

#==============================================================================================================
# xlnt-community
#==============================================================================================================
set(STATIC OFF CACHE BOOL "Build XLNT as SHARED library" FORCE)
add_subdirectory("kernel/tools/xlnt/")

#==============================================================================================================
# retrieve source code
#==============================================================================================================
list(APPEND PATHS "kernel/interfaces/" "kernel/openxlsx/" "kernel/xlnt/")
cmake_language(CALL COLLECT_SOURCE_CODE_AND_RESOURCES PATHS)

# create executable
add_library(XLSX ${SOURCES} ${HEADERS})
add_library(XLSX::XLSX ALIAS XLSX)
target_link_libraries(XLSX PUBLIC OpenXLSX::OpenXLSX xlnt)
target_include_directories(XLSX PUBLIC ${DIRECTORIES})

#==============================================================================================================
#installation rules
#==============================================================================================================
include(GNUInstallDirs)
install(TARGETS XLSX
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR})
