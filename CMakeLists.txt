cmake_minimum_required(VERSION 3.16)

project(XLSXView LANGUAGES CXX)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

#==============================================================================================================
# configure Qt
#==============================================================================================================
find_package(QT NAMES Qt6 Qt5 REQUIRED COMPONENTS Core Gui Widgets PdfWidgets Concurrent)
find_package(Qt${QT_VERSION_MAJOR} REQUIRED COMPONENTS Core Gui Widgets PdfWidgets Concurrent)

#==============================================================================================================
# include CMake macro
#==============================================================================================================
include("cmake/collect_source_code_and_resources.cmake")
add_subdirectory("source/")

#==============================================================================================================
# retrieve source code
#==============================================================================================================
list(APPEND PATHS "source/gui")
cmake_language(CALL COLLECT_SOURCE_CODE_AND_RESOURCES PATHS)

# create executable
add_executable(XLSXView ${SOURCES} ${HEADERS})
target_include_directories(XLSXView PUBLIC ${DIRECTORIES})

target_link_libraries(
    XLSXView
    PUBLIC

    #backend
    XLSX::XLSX

    #frontend
    Qt::Core
    Qt::Gui
    Qt::Widgets
)

#find_package(QT NAMES Qt6 Qt5 REQUIRED COMPONENTS Core)
#find_package(Qt${QT_VERSION_MAJOR} REQUIRED COMPONENTS Core)

#target_link_libraries(XLSX Qt${QT_VERSION_MAJOR}::Core)

#include(GNUInstallDirs)
#install(TARGETS XLSX
#    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
#    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
#)
