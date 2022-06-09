#include <xlnt/xlnt.hpp>
#include <emscripten/bind.h>
#include "./utils.h"

using namespace emscripten;


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
