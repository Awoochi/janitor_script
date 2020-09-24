
require 'roo'
require 'fileutils'

def find_files(doc)
  old_name = doc.sheet(0).column(3).drop(1)
  new_name = doc.sheet(0).column(4).drop(1)
  current_files = []
  new_codes = []
  new_paths = []
  filenames = []

  i = 0
  while i < old_name.length do
    old_name_code = old_name[i].match(/(\S+)(\d{2})/) # parse cells with old filenames
    old_filename = old_name_code[1]
    old_name_code = old_name_code[2]

    cur_file = Dir["#{Dir.pwd}/**/#{old_filename}#{old_name_code[0]}#{old_name_code[1]}*"] # return current file
    new_name_code = new_name[i].match(/(\S+)(\d{2})/) # parse cells with new filenames
    new_name_code = new_name_code[2]
    new_path = "#{Dir.pwd}/#{new_name_code[0]}/#{new_name_code[1]}/" # New filepath

    current_files.push(cur_file)
    new_codes.push(new_name_code)
    new_paths.push(new_path)
    filenames.push(old_filename)
    i += 1
  end

  return current_files, new_paths, new_codes
end

def move_files(current_files, new_paths, new_codes)
  j = 0
  while j < current_files.length do
    cur_file = current_files[j]
    next_path = new_paths[j]
    next_name = new_codes[j]
    if !(current_files[j].empty?)
      FileUtils.mv(cur_file, "#{next_path}")
    end
    j += 1
  end
end

def rename_files(files, next_filepaths, next_codes)
  t = 0
  while t < files.length
    cur_file = files[t]
    next_path = next_filepaths[t]
    next_code = next_codes[t]
    modify_files(cur_file, next_path, next_code)
    t += 1
  end
end

def modify_files(file, next_filepath, next_code)
  file.each do |f|
    file_match = f.match(/\S+\/(\S+)(\d{2})(\S+)/)
    filename = file_match[1]
    file_res = file_match[3]
    puts "#{f} | #{next_filepath} | #{next_code}"
    puts ". . ."
    File.rename(f, "#{next_filepath}#{filename}#{next_code}#{file_res}")
  end
end

xl_doc = Roo::Spreadsheet.open('./filenames.xlsx')
xl_doc = Roo::Excelx.new('./filenames.xlsx')

puts '-----FILES MODIFICATION STARTED-----'
result = find_files(xl_doc)
move_files(result[0], result[1], result[2])
next_result = find_files(xl_doc)
rename_files(next_result[0], next_result[1], next_result[2])
puts "-----MODIFICATION COMPLETED-----"
puts "Press ENTER to close the program..."
waiting = gets
