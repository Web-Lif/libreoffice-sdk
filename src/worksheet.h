#include <xlnt/xlnt.hpp>
#include <emscripten.h>
#include <emscripten/bind.h>

using namespace emscripten;

EMSCRIPTEN_BINDINGS(worksheet) {
  class_<xlnt::worksheet>("Worksheet")
    .function("garbage_collect", &xlnt::worksheet::garbage_collect)
    .function("cell", optional_override([](
        xlnt::worksheet& worksheet,
        uint32_t col,
        uint32_t row,
        emscripten::val action
    ){
        xlnt::cell cell = worksheet.cell(col, row);
        auto* cell_ptr = &cell;
        return action(cell_ptr);
    }));
}
