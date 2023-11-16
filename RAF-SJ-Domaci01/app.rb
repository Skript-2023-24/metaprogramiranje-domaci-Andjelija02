require "google_drive"

def from_camel_case(str)
  str.to_s.gsub(/([A-Z])/, ' \1').strip.downcase
end

class Table
  def initialize()
    session = GoogleDrive::Session.from_service_account_key("svc.json")
    @ws = session.spreadsheet_by_key("1WNmahbNbEiUGIVkTKcuYKbN2VczbibCri4WdWF_a4A0").worksheets[0]
    @start_col = 0
    @start_row = 0
    find_header
  end

  def find_header
    t = table
    t.each_with_index do |row, row_index|
      row.each_with_index do |cell, index|
        if t[row_index][index] != ""
          @start_col = index
          @start_row = row_index
          break
        end
      end
    end
  end

  def table
    return @ws.rows
  end

  def update_cell(row, col, new_value)
    @ws[row, col] = new_value
    @ws.save
  end

  def row(index)
    table[index]
  end

  def column(index)
    list = []
    table.each do |row|
      list << row[index]
    end
    list
  end

  def cells
    cells = []
    table().each do |row|
      cells += row
    end
    cells
  end

  def [](key)
    list = []
    r = row(start_row)
    r.each_with_index do |cell, index|
      if cell == key
        list = column(index)
        break
      end
    end
    list
  end
#Update the value
  def []=(key, value)
    column_name = key.to_s.capitalize
    index = row(start_row).index(column_name)

    if index
      # Find the corresponding column in the table
      column_values = table.transpose[index]
      # Update the value at the specified index in the column
      column_values[start_row] = value
      # Update the worksheet with the modified column
      column_values.each_with_index do |val, row_index|
        @ws[start_row + row_index, index + @start_col] = val
      end
      @ws.save
    end
  end


  def method_missing(method_name, *args, &block)
    if method_name.to_s.end_with?('=')
      # If the method ends with '=', treat it as a setter
      set_column_value(method_name.to_s.chomp('=').to_sym, args.first)
    else
      column_name = from_camel_case(method_name)
      row(start_row).each_with_index do |cell, index|
        if cell.downcase == column_name
          return NewArray.new(column(index))
        end
      end
      super  # call the original method_missing
  end
  end
end

def each(&block)
  table.each(&block)
end

def remove_total_subtotal_rows
  @ws.rows.reject! { |row| total_or_subtotal_row?(row) }
  @ws.save
end

def +(other_table)
  # Implement the logic to add two tables
  # Make sure headers are the same
  raise 'Tables have different headers' unless table_headers == other_table.table_headers

  # Add rows from the other table to the current table
  @ws.rows += other_table.table
  @ws.save

  self
end


def -(other_table)
  # Implement the logic to subtract one table from another
  # Make sure headers are the same
  raise 'Tables have different headers' unless table_headers == other_table.table_headers
  # Remove rows from the other table from the current table
  @ws.rows.reject! { |row| other_table.table.include?(row) }
  @ws.save
  self
end

  def table_headers
    table[start_row]
  end
private

def total_or_subtotal_row?(row)
  row.any? { |cell| cell.downcase.include?('total') || cell.downcase.include?('subtotal') }
end
def set_column_value(column_name, value)
  index = row(start_row).index(column_name.to_s.capitalize)

  if index
    # Find the corresponding column in the table
    column_values = table.transpose[index]

    # Update the value at the specified index in the column
    column_values[start_row] = value

    # Update the worksheet with the modified column
    column_values.each_with_index do |val, row_index|
      @ws[start_row + row_index, index + @start_col] = val
    end

    @ws.save
  end
end

class NewArray < Array
  def initialize(*args)
    super(args.flatten)
  end

  def sum
    inject(:+)
  end

  def avg
    sum.to_f / size
  end

  def method_missing(method_name, *args, &block)
    row_name = from_camel_case(method_name)
    each_with_index do |cell, index|
      return index if cell.downcase == row_name 
    end
    super
  end
end


t = Table.new()
print t.table()
print "\n"
# print t.row(1)
# print "\n"
# print t.cells()
# print "\n"
# print t["Prva Kolona"][1] = "Test"
# print t.prvaKolona[1]
# print t.drugaKolona
print t.prvaKolona.Test