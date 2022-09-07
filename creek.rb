require 'creek'
require 'write_xlsx'
# require 'ocra'

puts "START READ FILES"
workbook = Creek::Book.new('./Contact Rikai.xlsx')
puts "END READ FILES"

old_data = workbook.sheets[0]
new_data = workbook.sheets[1]

array_old_data_url = []
array_new_data_value = []


puts "start data_url_1 "
old_data.rows.each_with_index do |row, index|
    next if row.values[3].nil? 
    array_old_data_url << row.values[3]
end
puts "end data_url_1"


puts "start data 2"
new_data.rows.each_with_index do |row, index|
    next if row.values[3].nil? 
    array_new_data_value << row.values
end
puts "end data 2"


puts "start data_url_2"
array_new_data_url = []
array_new_data_value.each_with_index do |row, index|
    next if row[3].nil?
    array_new_data_url << row[3]
end
puts "end data_url_2"


data_url = (array_old_data_url | array_new_data_url)
data_url_double = (array_old_data_url & array_new_data_url)
data_url_undouble = data_url - (data_url_double)
data_url_new = data_url_undouble & array_new_data_url

p " ------------------------"
puts "START"
workbooks = WriteXLSX.new('demo1.xlsx')
worksheet = workbooks.add_worksheet


worksheet.write(0, 0 , array_new_data_value[0][0].to_s)
worksheet.write(0, 1 , array_new_data_value[0][1].to_s)
worksheet.write(0, 2 , array_new_data_value[0][2].to_s)
worksheet.write(0, 3 , array_new_data_value[0][3].to_s)
worksheet.write(0, 4 , array_new_data_value[0][4].to_s)
worksheet.write(0, 5 , array_new_data_value[0][5].to_s)

index = 1
array_new_data_value.each do |row|
    data_url_new_value = (row & data_url_new)
    next if data_url_new_value.nil?
    if  data_url_new_value.length > 0 
        worksheet.write(index, 0 , row[0])
        worksheet.write(index, 1 , row[1])
        worksheet.write(index, 2 , row[2])
        worksheet.write(index, 3 , row[3])
        index += 1;
    end
end

puts "END"
workbooks.close
