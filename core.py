
def sheet2cells(sheet):
  return map(lambda i: map(lambda j: j.value, i), sheet.get_rows())

def add_sheets(s0, s1):
  return list(map(add_rows, s0, s1))

def add_rows(r0, r1):
  return list(map(add_cells, r0, r1))

def add_cells(c0, c1):
  if type(c0) == float and type(c1) == float:
    return c0 + c1
  else:
    return c1
