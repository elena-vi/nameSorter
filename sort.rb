require 'rubyXL' # Assuming rubygems is already required

workbook = RubyXL::Parser.parse("Results.xlsx")  
worksheet = workbook[0]

names = Array.new
chest = Array.new
skin  = Array.new
champ = Array.new 


23.times do |c|
	row = [ (worksheet.sheet_data[c][1].value), (worksheet.sheet_data[c][2].value), (worksheet.sheet_data[c][3].value)]
	if row[1] == "Mystery Chest*"
		chest.push(row)
	end
	if row[1] == "Mystery Skin"
		skin.push(row)
	end
	if row[1] == "Mystery Champ"
		champ.push(row)
	end
	names.push(row)
end

puts names.length
puts "CHEST: #{chest.length}"

puts "SKIN: #{skin.length}"

skinPairs = Array.new


skin = skin.shuffle

while skin.length != 0
	one = skin.pop
	two = skin.pop
	skinPairs.push([one, two])
	skin = skin.shuffle
end

skinPairs.each_with_index do |val, index| 
	puts "#{val}" 
end

chestPairs = Array.new

chest = chest.shuffle

while chest.length != 0
	one = chest.pop
	two = chest.pop
	chestPairs.push([one, two])
	chest = chest.shuffle
end

chestPairs.each_with_index do |val, index| 
	puts "#{val}" 
end