cmake_minimum_required(VERSION 3.29.0)

project(sparse_package VERSION 0.1.0)

set(PACKAGE_NAME package)

set(PACKAGE_DIR ${CMAKE_CURRENT_BINARY_DIR}/package)

set(SIGNING_CERT_PATH "" CACHE FILEPATH "")

set(SIGNING_CERT_PASSWORD "" CACHE STRING "")

set(ROOT_CERT_PATH "" CACHE STRING "")

set(MSIX_OUTPUT ${CMAKE_CURRENT_BINARY_DIR}/${PACKAGE_NAME}_unsigned.msix)

set(MSIX_OUTPUT_SIGNED ${CMAKE_CURRENT_BINARY_DIR}/${PACKAGE_NAME}.msix)

set(APPXMANIFEST_FILE ${CMAKE_CURRENT_SOURCE_DIR}/AppxManifest.xml)

set(MSIX_EXTERNAL_PATH ${CMAKE_CURRENT_BINARY_DIR}/external)

# Prepare external files

add_custom_target(copy_sublime_explorer_commands
                  COMMAND ${CMAKE_COMMAND} -E copy "$<TARGET_FILE:sublime_explorer_commands>" "${MSIX_EXTERNAL_PATH}/msix/$<TARGET_FILE_NAME:sublime_explorer_commands>"
                  DEPENDS sublime_explorer_commands)

add_custom_target(copy_resources
                  COMMAND ${CMAKE_COMMAND} -E copy_directory "${CMAKE_CURRENT_SOURCE_DIR}/resources" "${MSIX_EXTERNAL_PATH}")

add_custom_target(prepare_external_files
                  DEPENDS copy_sublime_explorer_commands copy_resources)

# Prepare appxmanifest

add_custom_target(copy_appxmanifest
                  COMMAND ${CMAKE_COMMAND} -E copy "${APPXMANIFEST_FILE}" "${PACKAGE_DIR}/AppxManifest.xml")

# Build package

add_custom_command(OUTPUT ${MSIX_OUTPUT}
                   COMMAND makeappx pack /o /d "${PACKAGE_DIR}" /nv /p "${MSIX_OUTPUT}"
                   DEPENDS copy_appxmanifest
                   VERBATIM)

add_custom_target(make_msix
                  DEPENDS ${MSIX_OUTPUT})

# Sign package

add_custom_command(OUTPUT ${MSIX_OUTPUT_SIGNED}
                   COMMAND ${CMAKE_COMMAND} -E copy "${MSIX_OUTPUT}" "${MSIX_OUTPUT_SIGNED}"
                   COMMAND signtool sign /fd SHA256 /a /f "${SIGNING_CERT_PATH}" /p "${SIGNING_CERT_PASSWORD}" "${MSIX_OUTPUT_SIGNED}"
                   DEPENDS ${MSIX_OUTPUT}
                   VERBATIM)
                   

add_custom_target(sign_msix
                  DEPENDS ${MSIX_OUTPUT_SIGNED})

add_custom_target(sublime_package ALL DEPENDS sign_msix prepare_external_files)

install(FILES ${MSIX_OUTPUT_SIGNED} DESTINATION .)
install(DIRECTORY ${MSIX_EXTERNAL_PATH} DESTINATION .)

if(ROOT_CERT_PATH)
    install(FILES ${ROOT_CERT_PATH} DESTINATION . RENAME root_ca.cer)
endif()
