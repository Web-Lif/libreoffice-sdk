#include <xlnt/xlnt.hpp>
#include <emscripten/bind.h>
#include "./utils.h"

using namespace emscripten;

EMSCRIPTEN_BINDINGS(cell) {
    class_<xlnt::cell>("cell")
    .function("has_value", &xlnt::cell::has_value)
    // template <typename T> T value() const;
    .function("clear_value", &xlnt::cell::clear_value)
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
            return cell.value(std::nullptr_t{});
        }
        throwError("invalid parameter type.");
    }))
    // type data_type() const;
    // void data_type(type t);
    .function("garbage_collectible", &xlnt::cell::garbage_collectible)
    .function("is_date", &xlnt::cell::is_date)
    // cell_reference reference() const;
    .function("column", &xlnt::cell::column)
    .function("column_index", &xlnt::cell::column_index)
    .function("row", &xlnt::cell::row)
    // std::pair<int, int> anchor() const;
    // class hyperlink hyperlink() const;
    .function("hyperlink", static_cast<void(xlnt::cell::*)(const std::string&, const std::string&)>(&xlnt::cell::hyperlink))
    .function("hyperlink", static_cast<void(xlnt::cell::*)(xlnt::cell, const std::string&)>(&xlnt::cell::hyperlink))
    .function("hyperlink", static_cast<void(xlnt::cell::*)(xlnt::range, const std::string&)>(&xlnt::cell::hyperlink))
    .function("has_hyperlink", &xlnt::cell::has_hyperlink)
    // class alignment computed_alignment() const;
    // class border computed_border() const;
    // class fill computed_fill() const;
    // class font computed_font() const;
    // class number_format computed_number_format() const;
    // class protection computed_protection() const;
    .function("has_format", &xlnt::cell::has_format)
    // const class format format() const;
    // void format(const class format new_format);
    .function("clear_format", &xlnt::cell::clear_format)
    // class number_format number_format() const;
    // void number_format(const class number_format &format);
    // class font font() const;
    // void font(const class font &font_);
    // class fill fill() const;
    // void fill(const class fill &fill_);
    // class border border() const;
    // void border(const class border &border_);
    // class alignment alignment() const;
    // void alignment(const class alignment &alignment_);
    // class protection protection() const;
    // void protection(const class protection &protection_);
    .function("has_style", &xlnt::cell::has_style)
    // class style style();
    // const class style style() const;
    // void style(const class style &new_style);
    .function("style", static_cast<void(xlnt::cell::*)(const std::string&)>(&xlnt::cell::style))
    .function("clear_style", &xlnt::cell::clear_style)
    .function("formula", static_cast<std::string(xlnt::cell::*)() const>(&xlnt::cell::formula))
    .function("formula", static_cast<void(xlnt::cell::*)(const std::string&)>(&xlnt::cell::formula))
    .function("clear_formula", &xlnt::cell::clear_formula)
    .function("has_formula", &xlnt::cell::has_formula)
    .function("to_string", &xlnt::cell::to_string)
    .function("is_merged", &xlnt::cell::is_merged)
    .function("merged", &xlnt::cell::merged)
    .function("phonetics_visible", &xlnt::cell::phonetics_visible)
    .function("show_phonetics", &xlnt::cell::show_phonetics)
    // std::string error() const;
    // void error(const std::string &error);
    // cell offset(int column, int row);
    // class worksheet worksheet();
    // const class worksheet worksheet() const;
    // class workbook &workbook();
    // const class workbook &workbook() const;
    // calendar base_date() const;
    .function("check_string", &xlnt::cell::check_string)
    .function("has_comment", &xlnt::cell::has_comment)
    .function("clear_comment", &xlnt::cell::clear_comment)
    // class comment comment();
    // void comment(const std::string &text, const std::string &author = "Microsoft Office User");
    // void comment(const std::string &comment_text, const class font &comment_font, const std::string &author = "Microsoft Office User");
    // void comment(const class comment &new_comment);
    .function("width", &xlnt::cell::width)
    .function("height", &xlnt::cell::height)
    ;
}
