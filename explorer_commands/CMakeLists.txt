cmake_minimum_required(VERSION 3.29.0)

project(sublime_explorer_commands VERSION 0.1.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)

file(GLOB SOURCES ${CMAKE_CURRENT_SOURCE_DIR}/src/*)

set(DEF_FILE ${CMAKE_CURRENT_SOURCE_DIR}/${PROJECT_NAME}.def)

add_library(${PROJECT_NAME} SHARED ${SOURCES} ${DEF_FILE})

target_include_directories(${PROJECT_NAME} PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}/inc)

find_package(wil CONFIG REQUIRED)
target_link_libraries(${PROJECT_NAME} PRIVATE WIL::WIL)
target_link_libraries(${PROJECT_NAME} PRIVATE Shlwapi RuntimeObject)
install(TARGETS ${PROJECT_NAME} DESTINATION .)

if(CMAKE_EXPORT_COMPILE_COMMANDS)
	add_custom_target(compile_commands ALL COMMAND ${CMAKE_COMMAND} -E copy ${CMAKE_BINARY_DIR}/compile_commands.json ${CMAKE_SOURCE_DIR}/compile_commands.json DEPENDS ${CMAKE_BINARY_DIR}/compile_commands.json)
endif()