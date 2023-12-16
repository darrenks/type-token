WINDOW=10
N=100000
def calc_token_stats(output, row, cell)
  id = cell.lines[0]
  s = cell.lines[1..-1].join
  tokens = s.scan(/\p{L}+/).map(&:downcase)
  return if tokens.empty?
  output.add_cell(row,0,id)
  output.add_cell(row,1,tokens.size)
  output.add_cell(row,2,tokens.uniq.size)
  output.add_cell(row,3,tokens.uniq.size * 1.0/tokens.size)
  hist = tokens.group_by{|v|v}
  counts = hist.map{|k,v|v.size}

  x=0
  excess=tokens.size-WINDOW+1
  excess.times{|i|
    x+=tokens[i,WINDOW].uniq.size
  }
  x/=1.0*WINDOW*excess

  output.add_cell(row,4,x)

  x=0
  N.times{
    x+=tokens.sample(WINDOW).uniq.size
  }
  output.add_cell(row,5,x*1.0/N/WINDOW)
end

require 'rubyXL'

workbook = RubyXL::Parser.parse ($*[0] || raise("usage: ruby T-Lex.rb filename.xlsx"))
worksheets = workbook.worksheets
puts "Found #{worksheets.count} worksheets"

output = RubyXL::Workbook.new
output_sheet = output.add_worksheet('Sheet1')
output_sheet.add_cell(0,0,"id")
output_sheet.add_cell(0,1,"token count")
output_sheet.add_cell(0,2,"uniq count")
output_sheet.add_cell(0,3,"type token ratio")
output_sheet.add_cell(0,4,"moving avg window=#{WINDOW}")
output_sheet.add_cell(0,5,"true avg window=#{WINDOW}")

worksheets.each do |worksheet|
  puts "Reading: #{worksheet.sheet_name}"
  num_rows = 0
  worksheet.each do |row|
    num_rows += 1
    row_cells = row.cells.map{ |cell| calc_token_stats( output_sheet, num_rows, cell.value) }
  end
  puts "Read #{num_rows} rows"
end

output.write("out.xlsx")

# idea for when  have second set of data, use custom window for each student = min of their 2 samples

# also use https://github.com/jennafrens/lexical_diversity for jess

