# PRODSCH
Creates excel sheet after running PRODSCH in AppWorx

# REGEX explanations:

 __\s*__ zero or more white space characters
 
 __\s+__ one or more white space characters
 
 __\S+__ one or more non-whitespace characters
 
 __\w+__ one or more word characters (days of the week and month name)
 
 __\d+__ is the day (in numbers)
 
 __\d{4}__ is the year (2025)
 
 __\d{2}:\d{2}__ is the time in HH:MM (09:45)
 
 __?__ makes the group optional. Will match if it exists and will match if it is absent


 ## Parenthesis sets the Capture Groups:

  __(\S+)__ first Capture Group - name of chain or module

  __(\w+\s+\w+\s+\d+\s+\d{4}\s+\d{2}:\d{2})__ second Capture Group - date and time


 ## Capture these patterns:

  QUEUED       {Chain:OIT_C_ADASTRA_XFER} Thu Jan 30 2025 09:45 
 
  {Module:OIT_M_GZPEMAL} Thu Jan 30 2025 10:39 EST5EDT (GMT-5.0) (Dls)

---------------------------------------------------------------------------------------

# Tuple Deconstruction:

 var (sheet, row1, day) = prefixToSheet[prefix]; 
 
 Tuple deconstruction allows you to store multiple values in a single object, and you can extract them individually.
 
 prefixToSheet[prefix] accesses the value (the tuple (IXLWorksheet sheet, int row, string day)) associated with the current prefix in the dictionary.
 
 Ex: if prefix is OIT, then prefixToSheet["OIT'] might return a tuple like (OIT, 2, Wed)
