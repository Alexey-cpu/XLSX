cmake_minimum_required(VERSION 3.16)

project(XLSX LANGUAGES CXX)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

#==============================================================================================================
# include CMake macro
#==============================================================================================================
include("cmake/collect_source_code_and_resources.cmake")

#==============================================================================================================
# OpenXLSX
#==============================================================================================================
set(OPENXLSX_LIBRARY_TYPE "SHARED" CACHE STRING "Set the library type to SHARED or STATIC")
set(OPENXLSX_ENABLE_LTO OFF CACHE BOOL "Enables LTO for OpenXLSX" FORCE)
add_subdirectory("tools/OpenXLSX/OpenXLSX")

#==============================================================================================================
# xlnt-community
#==============================================================================================================
set(STATIC OFF CACHE BOOL "Build XLNT as SHARED library" FORCE)
add_subdirectory("tools/xlnt/")

#==============================================================================================================
# retrieve source code
#==============================================================================================================
list(APPEND PATHS "source/")
cmake_language(CALL COLLECT_SOURCE_CODE_AND_RESOURCES PATHS)

# create executable
add_executable(XLSX ${SOURCES} ${HEADERS}
    source/XLSX.h

    source/XLSX_OpenXLSX.h)

target_link_libraries(XLSX
    PUBLIC
    OpenXLSX::OpenXLSX
    xlnt
)

#find_package(QT NAMES Qt6 Qt5 REQUIRED COMPONENTS Core)
#find_package(Qt${QT_VERSION_MAJOR} REQUIRED COMPONENTS Core)

#target_link_libraries(XLSX Qt${QT_VERSION_MAJOR}::Core)

#include(GNUInstallDirs)
#install(TARGETS XLSX
#    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
#    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
#)
