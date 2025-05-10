cmake_minimum_required(VERSION 3.16)

project(XLSXView LANGUAGES CXX)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

message("------------------------------------------------------------------")
message("Short build info")
message("------------------------------------------------------------------")
message("Building project '${CMAKE_PROJECT_NAME}'")
message("Build type '${CMAKE_BUILD_TYPE}'")
message("------------------------------------------------------------------")

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
# configure common installation options
#==============================================================================================================
# Here we show CMake where we want all packages to be installed
set(CMAKE_INSTALL_PREFIX ${CMAKE_CURRENT_LIST_DIR}/dist/XLSXView/${CMAKE_SYSTEM_NAME} CACHE PATH "Force path to set CMAKE_INSTALL_PREFIX" FORCE)
set(CMAKE_WINDOWS_EXPORT_ALL_SYMBOLS ON CACHE BOOL "Force CMAKE_WINDOWS_EXPORT_ALL_SYMBOLS to TRUE" FORCE)
set(BUILD_SHARED_LIBS ON CACHE BOOL "Force BUILD_SHARED_LIBS" FORCE)

# As all our libraries are shared we need to show runtime executable where to look for .so files
# to be able to run out executable from the local installation folder
set(CMAKE_MACOSX_BUNDLE CACHE BOOL "Default CMAKE_MACOSX_BUNDLE to ON" TRUE)
set(CMAKE_MACOSX_RPATH CACHE BOOL "Default CMAKE_MACOSX_RPATH to ON" TRUE)
set(CMAKE_INSTALL_RPATH_USE_LINK_PATH TRUE CACHE BOOL "Force set CMAKE_INSTALL_RPATH_USE_LINK_PATH to TRUE" FORCE)
set(CMAKE_INSTALL_RPATH "${CMAKE_INSTALL_PREFIX}/lib" CACHE PATH "Force set CMAKE_INSTALL_PREFIX" FORCE)

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

#==============================================================================================================
# install
#==============================================================================================================

# configure installation
include(GNUInstallDirs)

install(TARGETS XLSXView
    BUNDLE DESTINATION ${CMAKE_INSTALL_PREFIX}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    INCLUDES DESTINATION ${CMAKE_INSTALL_INCLUDEDIR})

if(APPLE)

    install(CODE [[
        include(BundleUtilities)

        # Note Mac specific extension .app
        set(APPS ${CMAKE_INSTALL_PREFIX}/XLSXView.app)

        # Directories to look for dependencies
        list(APPEND DIRS ${CMAKE_INSTALL_PREFIX}/lib ${CMAKE_PREFIX_PATH})

        message(STATUS "APPS=${APPS}")
        message(STATUS "DIRS=${DIRS}")

        # NOTE: No need `${DIRS}` when already setting rpath in exectuable in install by `set(CMAKE_INSTALL_RPATH "${CMAKE_INSTALL_PREFIX}/lib")`
        fixup_bundle(${APPS} "" ${DIRS})

        # clean-up the folder
        file(REMOVE_RECURSE ${CMAKE_INSTALL_PREFIX}/lib ${CMAKE_INSTALL_PREFIX}/include ${CMAKE_INSTALL_PREFIX}/share)
    ]])

else()

endif()

# generate deploy script only on Windows and MacOS
if(WIN32 OR APPLE)

qt_generate_deploy_app_script(
    TARGET XLSXView
    OUTPUT_SCRIPT applicationDeployScript
    NO_UNSUPPORTED_PLATFORM_ERROR
)

install(SCRIPT ${applicationDeployScript})

# Package management
include(CPack)

endif()
