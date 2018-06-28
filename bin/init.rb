# Thanks to:
# Javaでexcelを読み書きするサンプル(Apache POIの使い方まとめ) | ぱーくん plus idea
# https://web.plus-idea.net/2016/01/how_to_read_excel_java_apache_poi/

require "csv"
require "bundler"
Bundler.require

require "jbundler"
java_import "org.apache.poi.ss.usermodel.Cell"
java_import "org.apache.poi.hssf.usermodel.HSSFWorkbook"

def get_cell_value(cell)
  case cell.cell_type
  when Cell.CELL_TYPE_BLANK
    nil
  when Cell.CELL_TYPE_BOOLEAN
    cell.boolean_cell_value?
  when Cell.CELL_TYPE_ERROR
    cell.error_cell_value
  when Cell.CELL_TYPE_FORMULA
    cell.cell_formula
  when Cell.CELL_TYPE_NUMERIC
    cell.numeric_cell_value
  when Cell.CELL_TYPE_STRING
    cell.string_cell_value
  else
    raise "Unknown cell type: #{cell.cell_type.inspect}"
  end
end

csv = CSV.new($stdout)

wb = HSSFWorkbook.new(java.lang.System.in)
wb.each do |sheet|
  sheet.each do |row|
    csv << row.map do |cell|
      get_cell_value(cell)
    end
  end
end
