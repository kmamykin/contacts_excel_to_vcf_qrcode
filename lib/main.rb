require 'spreadsheet'

FIRST_NAME = 0
LAST_NAME = 1
def process_row(row, index)
  return if row_empty?(row)
  puts "#{index} #{row[FIRST_NAME]} #{row[LAST_NAME]}" 
end

def row_empty?(row)
  value_blank?(row[FIRST_NAME]) && value_blank?(row[LAST_NAME])
end

def value_blank?(value)
  value.nil? or value.empty?
end

file = ARGV[0]
book = Spreadsheet.open file
sheet= book.worksheet(0)
puts "Skipping header #{sheet.row(0)}"
(1..sheet.row_count).to_a.each do |i|
  r = sheet.row(i)
  process_row(r, i)
end
