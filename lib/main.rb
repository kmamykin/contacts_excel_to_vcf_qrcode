require 'cgi'
require 'open-uri'
require 'spreadsheet'
require 'vpim/vcard'

FIRST_NAME = 0
LAST_NAME = 1

file = ARGV[0]

def cleanup_files
  puts Dir.glob("data/*.png") { |f| File.delete f }
  puts Dir.glob("data/*.vcf") { |f| File.delete f }
end

def generate_vcard(row)
  Vpim::Vcard::Maker.make2 do |maker|
    maker.add_name do |name|
#      name.prefix = 'Dr.'
      name.given = row[FIRST_NAME]
      name.family = row[LAST_NAME]
    end

    maker.org = "Company X"

    maker.title = "title awesome"

    maker.add_addr do |addr|
      addr.location = 'WORK'
      addr.street = '12 Last Row, 13th Section'
      addr.locality = 'City of Lost Children'
      addr.country = 'Cinema'
    end

#    maker.nickname = "The Good Doctor"

    maker.add_tel('416-123-5555') do |tel|
      tel.preferred = true
      tel.location = 'WORK'
    end
    maker.add_tel('416-123-3333') do |tel|
      tel.location = 'CELL'
    end

    maker.add_email('drdeath@work.com') do |email|
      email.location = 'WORK'
    end
  end
end

def qr_code_url(data)
  size = 250
  "http://chart.apis.google.com/chart?cht=qr&chl=#{CGI.escape(data.to_s)}&choe=UTF-8&chld=M&chs=#{size}x#{size}"
end

def save_vcard(vcard, index)
  open("data/contact#{index}.vcf", 'wb') do |file|
    file << vcard.to_s
  end
end

def save_vcard_as_qrcode(vcard, index)
  url = qr_code_url(vcard)
  open("data/qr#{index}.png", 'wb') do |file|
    file << open(url).read
  end
end

def process_row(row, index)
  return if row_empty?(row)
  puts "Processing #{index} #{row[FIRST_NAME]} #{row[LAST_NAME]}"
  card = generate_vcard(row)
  save_vcard(card, index)
  save_vcard_as_qrcode(card, index)
end

def row_empty?(row)
  value_blank?(row[FIRST_NAME]) && value_blank?(row[LAST_NAME])
end

def value_blank?(value)
  value.nil? or value.empty?
end

book = Spreadsheet.open file
sheet= book.worksheet(0)
puts "Skipping header #{sheet.row(0)}"
cleanup_files
(1..sheet.row_count).to_a.each do |i|
  r = sheet.row(i)
  process_row(r, i)
end
