cmake_minimum_required(VERSION 3.29.0)

project(explorer_commands VERSION 0.1.0 LANGUAGES CXX)

add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/explorer_commands EXCLUDE_FROM_ALL)

add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/sparse_package)