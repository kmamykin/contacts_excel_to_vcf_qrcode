require 'spreadsheet'
require 'vpim/vcard'

FIRST_NAME = 0
LAST_NAME = 1

def generate_vcard(row)
  card = Vpim::Vcard::Maker.make2 do |maker|
    maker.add_name do |name|
      name.prefix = 'Dr.'
      name.given = row[FIRST_NAME]
      name.family = row[LAST_NAME]
    end

    maker.add_addr do |addr|
      addr.preferred = true
      addr.location = 'work'
      addr.street = '12 Last Row, 13th Section'
      addr.locality = 'City of Lost Children'
      addr.country = 'Cinema'
    end

    maker.add_addr do |addr|
      addr.location = ['home', 'zoo']
      addr.delivery = ['snail', 'stork', 'camel']
      addr.street = '12 Last Row, 13th Section'
      addr.locality = 'City of Lost Children'
      addr.country = 'Cinema'
    end

    maker.nickname = "The Good Doctor"

    maker.birthday = Date.today

    maker.add_photo do |photo|
      photo.link = 'http://example.com/image.png'
    end

    maker.add_photo do |photo|
      photo.image = "File.open('drdeath.jpg').read # a fake string, real data is too large :-)"
      photo.type = 'jpeg'
    end

    maker.add_tel('416 123 1111')

    maker.add_tel('416 123 2222') { |t| t.location = 'home'; t.preferred = true }

    maker.add_impp('joe') do |impp|
      impp.preferred = 'yes'
      impp.location = 'mobile'
    end

    maker.add_x_aim('example') do |xaim|
      xaim.location = 'row12'
    end

    maker.add_tel('416-123-3333') do |tel|
      tel.location = 'work'
      tel.capability = 'fax'
    end

    maker.add_email('drdeath@work.com') { |e| e.location = 'work' }

    maker.add_email('drdeath@home.net') { |e| e.preferred = 'yes' }

  end
  card
end

def process_row(row, index)
  return if row_empty?(row)
  puts "Processing #{index} #{row[FIRST_NAME]} #{row[LAST_NAME]}"
  card = generate_vcard(row)
  puts card.to_s
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
