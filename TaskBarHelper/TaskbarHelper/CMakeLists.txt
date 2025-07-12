cmake_minimum_required(VERSION 3.10)

project(TaskbarHelper LANGUAGES CXX)

set(TASKBAR_HELPER_SRCDIR "src")
set(TASKBAR_HELPER_INCDIR "include")

add_library(TaskbarHelper STATIC
    ${TASKBAR_HELPER_SRCDIR}/TaskbarHelper.cpp
)

target_compile_definitions(TaskbarHelper PRIVATE UNICODE)

target_compile_options(TaskbarHelper PRIVATE
    /EHsc
)

target_include_directories(TaskbarHelper PRIVATE
    ${TASKBAR_HELPER_INCDIR}
)

set_target_properties(TaskbarHelper PROPERTIES
    ARCHIVE_OUTPUT_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
)