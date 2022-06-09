cmake_minimum_required(VERSION 3.0.0)
project(xlnt-es VERSION 0.1.0)


include(CTest)
enable_testing()

add_executable(
    ${PROJECT_NAME}
    src/index.cpp
    src/worksheet.h
    src/workbook.h
    src/cell.h
    src/utils.h
)

add_subdirectory(third-party/xlnt/source)

set(CPACK_PROJECT_NAME ${PROJECT_NAME})
set(CPACK_PROJECT_VERSION ${PROJECT_VERSION})

target_link_options(
    ${PROJECT_NAME} PRIVATE
)

target_link_libraries(
    ${PROJECT_NAME}
    xlnt
    -lembind
)

include(CPack)
