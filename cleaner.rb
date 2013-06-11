#!/usr/bin/env ruby

require 'madison'
require 'simple-spreadsheet'
require 'geokit'
require 'chronic'
require 'roo'

s = SimpleSpreadsheet::Workbook.read(ARGV[0])
puts ARGV
# s = SimpleSpreadsheet::Workbook.read("backup_regis.csv")

book = Spreadsheet::Workbook.new
write_sheet = book.create_worksheet


m = Madison.new

s.first_row.upto(s.last_row) do |line|
  next if line==s.first_row
  geo = GeoKit::Geocoders::MultiGeocoder.multi_geocoder(s.cell(line,27))
  zip = geo.zip.strip if !geo.zip.nil? && !geo.zip.empty?
  year = Chronic::parse(s.cell(line,34))
  join_date = Chronic::parse(s.cell(line,1)).strftime('%-m/%-d/%Y')
  state=s.cell(line,26)
  state = m.get_abbrev(s.cell(line,26)) unless m.get_abbrev(s.cell(line,26)).nil?

  write_sheet.row(line)[0] = s.cell(line,21).capitalize
  write_sheet.row(line)[1] = s.cell(line,20).capitalize
  write_sheet.row(line)[2] = s.cell(line,23)
  write_sheet.row(line)[3] = s.cell(line,24)
  write_sheet.row(line)[4] = s.cell(line,25)
  write_sheet.row(line)[5] = state
  write_sheet.row(line)[6] = zip
  write_sheet.row(line)[7] = s.cell(line,36)
  write_sheet.row(line)[8] = s.cell(line,34)
  write_sheet.row(line)[9] = year
  write_sheet.row(line)[10] = s.cell(line,29)
  write_sheet.row(line)[11] = s.cell(line,30)
  write_sheet.row(line)[12] = join_date
  write_sheet.row(line)[13] = s.cell(line,32)  

end

# book.write File.basename("backup_regis.csv") + ".xls"
book.write File.basename(ARGV[0]) + ".xls"

