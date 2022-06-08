#include <xlnt/xlnt.hpp>
#include <emscripten.h>
#include <emscripten/bind.h>

using namespace emscripten;

void throwError(const std::string& msg)
{
    EM_ASM_({ throw UTF8ToString($0); }, msg.c_str());
}


EMSCRIPTEN_BINDINGS(workbook) {
  class_<xlnt::workbook>("Workbook")
    .constructor<>()
    .function("active_sheet", &xlnt::workbook::active_sheet)
    .function("create_sheet",  optional_override([](
        xlnt::workbook& workbook,
        emscripten::val value,
        emscripten::val action
    ){
        auto type = value.typeOf().as<std::string>();
        auto valueType = value.typeOf().as<std::string>();
        if (type == "number" && valueType == "function") {
            xlnt::worksheet sheet = workbook.create_sheet(value.as<std::size_t>());
            auto* sheet_ptr = &sheet;
            return action(sheet_ptr);
        } else if (valueType == "function") {
            xlnt::worksheet sheet = workbook.create_sheet();
            auto* sheet_ptr = &sheet;
            return action(sheet_ptr); 
        }
        throwError("invalid parameter type.");
    }));
}

EMSCRIPTEN_BINDINGS(worksheet) {
  class_<xlnt::worksheet>("Worksheet")
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


EMSCRIPTEN_BINDINGS(cell) {
    class_<xlnt::cell>("cell")
    .function("set_value", emscripten::optional_override([](
        xlnt::cell& cell,
        emscripten::val value)
    {
        auto type = value.typeOf().as<std::string>();
        if (type == "number") {
            return cell.value(value.as<double>());
        } else if (type == "string") {
            return cell.value(value.as<std::string>());
        } else if (type == "boolean") {
            return cell.value(value.as<bool>());
        } else if (value.isNull() || value.isUndefined()) {
            return cell.value(std::nullptr_t{});
        }
        throwError("invalid parameter type.");
    }));
}

