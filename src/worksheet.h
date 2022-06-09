#include <xlnt/xlnt.hpp>
#include <emscripten.h>
#include <emscripten/bind.h>

using namespace emscripten;

EMSCRIPTEN_BINDINGS(worksheet) {
  class_<xlnt::worksheet>("Worksheet")
    .function("garbage_collect", &xlnt::worksheet::garbage_collect)
    .function("id", static_cast<std::size_t(xlnt::worksheet::*)() const>(&xlnt::worksheet::id))
    .function("id", static_cast<void(xlnt::worksheet::*)(std::size_t)>(&xlnt::worksheet::id))
    .function("title", static_cast<std::string(xlnt::worksheet::*)() const>(&xlnt::worksheet::title))
    .function("title", static_cast<void(xlnt::worksheet::*)(const std::string&)>(&xlnt::worksheet::title))
    // cell_reference frozen_panes() const;
    .function("freeze_panes", static_cast<void(xlnt::worksheet::*)(xlnt::cell)>(&xlnt::worksheet::freeze_panes))
    // void freeze_panes(const cell_reference &top_left_coordinate);
    .function("unfreeze_panes", &xlnt::worksheet::unfreeze_panes)
    .function("has_frozen_panes", &xlnt::worksheet::has_frozen_panes)
    // bool has_cell(const cell_reference &reference) const;
    // class cell cell(const cell_reference &reference);
    // const class cell cell(const cell_reference &reference) const;
    .function("cell", optional_override([](
        xlnt::worksheet& worksheet,
        uint32_t col,
        uint32_t row,
        emscripten::val action
    ){
        xlnt::cell cell = worksheet.cell(col, row);
        auto* cell_ptr = &cell;
        return action(cell_ptr);
    }))
    // const class cell cell(column_t column, row_t row) const;
    // class range range(const std::string &reference_string);
    // const class range range(const std::string &reference_string) const;
    // class range range(const range_reference &reference);
    // const class range range(const range_reference &reference) const;
    // class range rows(bool skip_null = true);
    // const class range rows(bool skip_null = true) const;
    // class range columns(bool skip_null = true);
    // const class range columns(bool skip_null = true) const;
    // void clear_cell(const cell_reference &ref);
    .function("clear_row",  &xlnt::worksheet::clear_row)
    .function("insert_rows",  &xlnt::worksheet::insert_rows)
    .function("insert_columns", &xlnt::worksheet::insert_columns)
    .function("delete_rows", &xlnt::worksheet::delete_rows)
    .function("delete_columns", &xlnt::worksheet::delete_columns)
    // xlnt::column_properties &column_properties(column_t column);
    // const xlnt::column_properties &column_properties(column_t column) const;
    .function("has_column_properties", &xlnt::worksheet::has_column_properties)
    // void add_column_properties(column_t column, const class column_properties &props);
    .function("column_width", &xlnt::worksheet::column_width)
    // xlnt::row_properties &row_properties(row_t row);
    // const xlnt::row_properties &row_properties(row_t row) const;
    .function("has_row_properties", &xlnt::worksheet::has_row_properties)
    // void add_row_properties(row_t row, const class row_properties &props);
    .function("row_height", &xlnt::worksheet::row_height)
    // cell_reference point_pos(int left, int top) const;
    .function("create_named_range", static_cast<void(xlnt::worksheet::*)(const std::string&, const std::string&)>(&xlnt::worksheet::create_named_range))
    .function("has_named_range", &xlnt::worksheet::has_named_range)
    // class range named_range(const std::string &name);
    // const class range named_range(const std::string &name) const;
    .function("remove_named_range", &xlnt::worksheet::remove_named_range)
    .function("lowest_row", &xlnt::worksheet::lowest_row)
    .function("lowest_row_or_props", &xlnt::worksheet::lowest_row_or_props)
    .function("highest_row", &xlnt::worksheet::highest_row)
    .function("highest_row_or_props", &xlnt::worksheet::highest_row_or_props)
    .function("next_row", &xlnt::worksheet::next_row)
    .function("lowest_column", &xlnt::worksheet::lowest_column)
    .function("lowest_column_or_props", &xlnt::worksheet::lowest_column_or_props)
    .function("highest_column", &xlnt::worksheet::highest_column)
    .function("highest_column_or_props", &xlnt::worksheet::highest_column_or_props)
    // range_reference calculate_dimension(bool skip_null=true) const;
    .function("merge_cells", static_cast<void(xlnt::worksheet::*)(const std::string&)>(&xlnt::worksheet::merge_cells))
    // void merge_cells(const range_reference &reference);
    .function("unmerge_cells", static_cast<void(xlnt::worksheet::*)(const std::string&)>(&xlnt::worksheet::unmerge_cells))
    // void unmerge_cells(const range_reference &reference);
    .function("merged_ranges", &xlnt::worksheet::merged_ranges)
    .function("has_page_setup", &xlnt::worksheet::has_page_setup)
    // xlnt::page_setup page_setup() const;
    // void page_setup(const struct page_setup &setup);
    .function("has_page_margins", &xlnt::worksheet::has_page_margins)
    // xlnt::page_margins page_margins() const;
    // void page_margins(const class page_margins &margins)
    // void auto_filter(const std::string &range_string);
    // void auto_filter(const xlnt::range &range);
    // void auto_filter(const range_reference &reference);
    .function("clear_auto_filter", &xlnt::worksheet::clear_auto_filter)
    .function("has_auto_filter", &xlnt::worksheet::has_auto_filter)
    .function("reserve", &xlnt::worksheet::reserve)
    .function("has_phonetic_properties", &xlnt::worksheet::has_phonetic_properties)
    // const phonetic_pr &phonetic_properties() const;
    // void phonetic_properties(const phonetic_pr &phonetic_props);
    .function("has_header_footer", &xlnt::worksheet::has_header_footer)
    // class header_footer header_footer() const;    
    // void header_footer(const class header_footer &new_header_footer);
    // xlnt::sheet_state sheet_state() const;
    // void sheet_state(xlnt::sheet_state state);
    .function("print_title_rows", static_cast<void(xlnt::worksheet::*)(xlnt::row_t, xlnt::row_t)>(&xlnt::worksheet::print_title_rows))
    .function("print_title_rows", static_cast<void(xlnt::worksheet::*)(xlnt::row_t)>(&xlnt::worksheet::print_title_rows))
    .function("print_title_cols", static_cast<void(xlnt::worksheet::*)(xlnt::column_t, xlnt::column_t)>(&xlnt::worksheet::print_title_cols))
    .function("print_title_cols", static_cast<void(xlnt::worksheet::*)(xlnt::column_t)>(&xlnt::worksheet::print_title_cols))
    .function("print_titles", &xlnt::worksheet::print_titles)
    .function("print_area", static_cast<void(xlnt::worksheet::*)(const std::string&)>(&xlnt::worksheet::print_area))
    // range_reference print_area() const;
    .function("has_view", &xlnt::worksheet::has_view)
    // sheet_view &view(std::size_t index = 0) const;
    // void add_view(const sheet_view &new_view);
    // void active_cell(const cell_reference &ref);
    .function("has_active_cell", &xlnt::worksheet::has_active_cell)
    // cell_reference active_cell() const;
    .function("clear_page_breaks", &xlnt::worksheet::clear_page_breaks)
    .function("page_break_rows", &xlnt::worksheet::page_break_rows)
    .function("page_break_at_row", &xlnt::worksheet::page_break_at_row)
    .function("page_break_columns", &xlnt::worksheet::page_break_columns)
    .function("page_break_at_column", &xlnt::worksheet::page_break_at_column)
    // xlnt::conditional_format conditional_format(const range_reference &ref, const condition &when);
    // xlnt::path path() const;
    // relationship referring_relationship() const;
    // sheet_format_properties format_properties() const;
    // void format_properties(const sheet_format_properties &properties);
    .function("has_drawing", &xlnt::worksheet::has_drawing)
    .function("is_empty", &xlnt::worksheet::is_empty)
    ;

}
