#include <xlnt/xlnt.hpp>
#include <emscripten.h>
#include <emscripten/bind.h>
#include "./utils.h"

using namespace emscripten;

EMSCRIPTEN_BINDINGS(workbook) {
  class_<xlnt::workbook>("Workbook")
    .constructor<>()
    .function("create_sheet", optional_override([](
        xlnt::workbook& workbook,
        emscripten::val action
    ){
        xlnt::worksheet sheet = workbook.create_sheet();
        auto* sheet_ptr = &sheet;
        return action(sheet_ptr);
    }))
    .function("create_sheet", optional_override([](
        xlnt::workbook& workbook,
        std::size_t index,
        emscripten::val action
    ){
        xlnt::worksheet sheet = workbook.create_sheet(index);
        auto* sheet_ptr = &sheet;
        return action(sheet_ptr);
    }))
    .function("copy_sheet", optional_override([](
        xlnt::workbook& workbook,
        xlnt::worksheet& worksheet,
        emscripten::val action
    ){
        xlnt::worksheet sheet = workbook.copy_sheet(worksheet);
        auto* sheet_ptr = &sheet;
        return action(sheet_ptr);
    }))
    .function("copy_sheet", optional_override([](
        xlnt::workbook& workbook,
        xlnt::worksheet& worksheet,
        std::size_t index,
        emscripten::val action
    ){
        xlnt::worksheet sheet = workbook.copy_sheet(worksheet, index);
        auto* sheet_ptr = &sheet;
        return action(sheet_ptr);
    }))
    .function("active_sheet", optional_override([](
        xlnt::workbook& workbook,
        emscripten::val action
    ){
        xlnt::worksheet sheet = workbook.active_sheet();
        auto* sheet_ptr = &sheet;
        return action(sheet_ptr);
    }))
    .function("sheet_by_title", optional_override([](
        xlnt::workbook& workbook,
        std::string title,
        emscripten::val action
    ){
        xlnt::worksheet sheet = workbook.sheet_by_title(title);
        auto* sheet_ptr = &sheet;
        return action(sheet_ptr);
    }))
    .function("sheet_by_index", optional_override([](
        xlnt::workbook& workbook,
        std::size_t index,
        emscripten::val action
    ){
        xlnt::worksheet sheet = workbook.sheet_by_index(index);
        auto* sheet_ptr = &sheet;
        return action(sheet_ptr);
    }))
    .function("sheet_by_id", optional_override([](
        xlnt::workbook& workbook,
        std::size_t id,
        emscripten::val action
    ){
        xlnt::worksheet sheet = workbook.sheet_by_id(id);
        auto* sheet_ptr = &sheet;
        return action(sheet_ptr);
    }))
    .function("sheet_hidden_by_index", &xlnt::workbook::sheet_hidden_by_index)
    .function("contains", &xlnt::workbook::contains)
    .function("index", &xlnt::workbook::index)
    .function("remove_sheet", &xlnt::workbook::remove_sheet)
    .function("clear", &xlnt::workbook::clear)
    .function("sheet_titles", &xlnt::workbook::sheet_titles)
    .function("sheet_count", &xlnt::workbook::sheet_count)
    .function("has_core_property", &xlnt::workbook::has_core_property)
    .function("core_properties", &xlnt::workbook::core_properties)
    // variant core_property(xlnt::core_property type) const;
    // void core_property(xlnt::core_property type, const variant &value);

    .function("has_extended_property", &xlnt::workbook::has_extended_property)
    .function("extended_properties", &xlnt::workbook::extended_properties)
    // variant extended_property(xlnt::extended_property type) const;
    // void extended_property(xlnt::extended_property type, const variant &value);
    // bool has_custom_property(const std::string &property_name) const;
    .function("custom_properties", &xlnt::workbook::custom_properties)
    // variant custom_property(const std::string &property_name) const;
    // void custom_property(const std::string &property_name, const variant &value);
    // calendar base_date() const;
    // void base_date(calendar base_date);
    .function("has_title", &xlnt::workbook::has_title)
    .function("title", static_cast<std::string(xlnt::workbook::*)() const>(&xlnt::workbook::title))
    .function("title", static_cast<void(xlnt::workbook::*)(const std::string&)>(&xlnt::workbook::title))
    .function("abs_path", &xlnt::workbook::abs_path)
    .function("arch_id_flags", &xlnt::workbook::arch_id_flags)
    // std::vector<xlnt::named_range> named_ranges() const;
    // void create_named_range(const std::string &name, worksheet worksheet, const range_reference &reference);
    // void create_named_range(const std::string &name, worksheet worksheet, const std::string &reference_string);
    .function("has_named_range", &xlnt::workbook::has_named_range)
    // class range named_range(const std::string &name);
    .function("remove_named_range", &xlnt::workbook::remove_named_range)
    // void save(std::vector<std::uint8_t> &data) const;
    // void save(std::vector<std::uint8_t> &data, const std::string &password) const;
    .function("save",
        static_cast<void(xlnt::workbook::*)(const std::string&) const>(
            &xlnt::workbook::save))
    .function("save",
        static_cast<void(xlnt::workbook::*)(const std::string&, const std::string&) const>(
            &xlnt::workbook::save));
}
