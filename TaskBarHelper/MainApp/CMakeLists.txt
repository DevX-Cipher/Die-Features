cmake_minimum_required(VERSION 3.10)

project(WindowsProject1 LANGUAGES CXX)

set(WINDOWS_PROJECT1_SRCDIR "${CMAKE_CURRENT_SOURCE_DIR}/src")
set(WINDOWS_PROJECT1_INCDIR "${CMAKE_CURRENT_SOURCE_DIR}/include")
set(WINDOWS_PROJECT1_RESDIR "${CMAKE_CURRENT_SOURCE_DIR}/res")

set(WINDOWS_PROJECT1_SOURCES
    ${WINDOWS_PROJECT1_SRCDIR}/WindowsProject1.cpp
    ${WINDOWS_PROJECT1_RESDIR}/WindowsProject1.rc
)

add_executable(WindowsProject1 ${WINDOWS_PROJECT1_SOURCES})

target_compile_definitions(WindowsProject1 PRIVATE UNICODE)
target_compile_options(WindowsProject1 PRIVATE /EHsc)

target_include_directories(WindowsProject1 PRIVATE
    ${WINDOWS_PROJECT1_INCDIR}
    ${CMAKE_SOURCE_DIR}/TaskbarHelper/include
)

target_link_libraries(WindowsProject1 PRIVATE TaskbarHelper runtimeobject)

set_target_properties(WindowsProject1 PROPERTIES
    WIN32_EXECUTABLE TRUE
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}/bin"
)
