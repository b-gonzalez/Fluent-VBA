# Regex file in python to parse and generate new dynamic tests that were originally written in mTests.

import re

def get_test_value_test_input(temp):
  # testValTestInputpattern = r"(?<=\()[\w|,|\s|\"|-]+(?=\))"
  # testValTestInputpattern = r"(?<=\()[\w|\.|\[|_|\]|,|\s\(]+(?=\))"
  # testValTestInputpattern = r"(?<=\()[\w|\.|\[|_|\]|,|\s\(]+(?=\))|\(\)"
  # testValTestInputpattern = r"(?<=\()[\w|\.|\[|_|\]|,|\s|\(|\"]+(?=\))|\(\)"
  # testValTestInputpattern = r"(?<=\().+(?=\).)|(?<=\().+(?=\)$)|\(\)"
  # testValTestInputpattern = r"(?<=\().+(?=\)\.)|(?<=\().+(?=\)$)|\(\)"
  testValTestInputpattern = r"(?<=\().+(?=\)\.)|(?<=\().+(?=\))|\(\)"

  testValTestInputList = re.findall(testValTestInputpattern, temp)

  value_input_dict = {}

  if testValTestInputList[0] != "()":
      value_input_dict['testingValue'] = testValTestInputList[0]

  if len(testValTestInputList) == 2:
    if not "VBA.[_HiddenModule].Array(" in testValTestInputList[1] and "," in testValTestInputList[1]:
      temp = []
      temp = testValTestInputList[1].split(",")

      for index, elem in enumerate(temp):
        vn = f"testingInput{index + 1}"
        value_input_dict[vn] = elem
    else:
      vn = f"testingInput"
      value_input_dict[vn] = testValTestInputList[1]

  return value_input_dict

def get_should_bool(temp):
  shouldPattern = r"(?<=\)\.)\w+(?=\.)"

  shouldMatch = re.search(shouldPattern, temp)

  if shouldMatch:
    b = None

    if shouldMatch.group(0) == "Should":
      b = True
    elif shouldMatch.group(0) == "ShouldNot":
      b = False

    return f'shouldMatch = {b}'

def get_funtion_name(temp, functionName=""):
  # functionNamePattern = r"(?<=\.)\w+(?=\([\w|,|\s]+\)$|$)"
  # functionNamePattern = r"(?<=\.)\w+(?=\([\w|,|\s|\"]+\)$|$)"
  # functionNamePattern = r"(?<=\.)\w+(?=\([\w|,|\s|\"|\.|!]+\)$|$)"
  # functionNamePattern = r"(?<=\.)\w+(?=\([\w|,|\s|\"|\.|!\(\)]+\)$|$)"
  # functionNamePattern = r"(?<=\.)\w+(?=\()"
  # functionNamePattern = r"(?<=\.\w{2}\.)\w+|(?<=\.\w{4}\.)\w+"
  if functionName == "":
    functionNamePattern = r"(?<=\.\w{2}\.)\w+|(?<=\.\w{5}\.)\w+|(?<=\.\w{4}\.)\w+|(?<=\.\w{7}\.)\w+|(?<=\.\w{6}\.)\w+|(?<=\.\w{9}\.)\w+"

    functionNameMatch = re.findall(functionNamePattern, temp)

    if functionNameMatch:
      # return f'functionName = "{functionNameMatch[-1]}"'
      return f'{functionNameMatch[-1]}'
  else:
    return functionName

def get_fluent_test(var_names, functionName = ""):
  temp = ""
  final = []

  assignment = "Set fluentInput"

  if functionName in ["Between","LengthBetween"] and len(var_names[1:]) == 2:
    final.append(f"testingValue:={var_names[0]}")
    final.append(f"lowerVal:={var_names[1]}")
    final.append(f"higherVal:={var_names[2]}")
  else:
    for vn in var_names:
      final.append(f"{vn}:={vn}")

  if len(var_names) == 0:
    temp = f"{assignment} = fluentTester(fluentInput, functionName, shouldMatch)"
  elif len(var_names) == 1:
      temp = f"{assignment} = fluentTester(fluentInput, functionName, shouldMatch, {var_names[0]}:={var_names[0]})"
  elif len(var_names) > 0:
      temp = f"{assignment} = fluentTester(fluentInput, functionName, shouldMatch, {", ".join(final)})"

  return temp

def is_test_line(temp):
  pattern = r"^\s*\w+\.\w+\s=\s"
  testLine = re.search(pattern, temp)

  return testLine != None

def get_indent(temp):
  pattern = r"^\s*"

  indentMatch = re.search(pattern, temp)

  if indentMatch:
    temp = indentMatch.group(0)

  return temp

def read_and_get_lines_from_file(filename):
  lines = []
  final = []

  with open(filename, "r") as file:
    lines = file.readlines()
    temp = ""

    for line in lines:
      if "AssertAndRaiseEvents" in line:
        temp = line.replace("testFluent,","fluentInput,")
        temp = temp.replace("AssertAndRaiseEvents","AssertAndRaiseEventsRefactor") 
      else:
        temp = line
      final.append(temp)

  return final

def write_list_to_file(str_list, file_name, functionName = ""):
  with open(file_name, "w") as file:
      for item in str_list:
        if is_test_line(item):
          indent = get_indent(item)
          value_input_dict = get_test_value_test_input(item)

          for key in value_input_dict.keys():
            if value_input_dict[key] == 'd' or value_input_dict[key] in ['col','col2','testFluent']:
              file.write(f"{indent}Set {key} = {value_input_dict[key]}" + "\n")
            else:
              file.write(f"{indent}{key} = {value_input_dict[key]}" + "\n")

          file.write(f"{indent}{get_should_bool(item)}" + "\n")

          tempfunctionName = ""
          
          if functionName == "":
            tempfunctionName = get_funtion_name(item)
            file.write(f'{indent}functionName = "{tempfunctionName}"' + "\n")
          else:
            tempfunctionName = functionName
            file.write(f'{indent}functionName = "{tempfunctionName}"' + "\n")

          file.write(f"{indent}{get_fluent_test(list(value_input_dict.keys()),tempfunctionName)}" + "\n") 
        else:
          file.write(f"{item}")  

lines = read_and_get_lines_from_file(r"C:\myDir\test_input.txt")

write_list_to_file(lines, r"C:\myDir\test_output.txt")

print("Finished!")