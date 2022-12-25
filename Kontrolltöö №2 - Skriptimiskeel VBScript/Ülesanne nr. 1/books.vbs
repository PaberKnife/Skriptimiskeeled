book = WScript.Arguments(0)
amount = WScript.Arguments(1)

set file = CreateObject("Scripting.FileSystemObject").OpenTextFile(book,1)
set dict = CreateObject("Scripting.Dictionary")

do until file.AtEndOfStream
    line = LCase(file.ReadLine())
    tokens = Split(line, " ")
    for each word in tokens
      if word = "don't" or word = "doesn't" then
        word = "do not"
      elseif word = "can't" then
        word = "cannot"
      elseif word = "an" then
        word = "a"
      elseif word = "me" or word = "mine" or word = "my" or word = "myself" then
        word = "i"
      elseif word = "him" or word = "his" or word = "himself" then
        word = "he"
      elseif word = "her" or word = "hers" or word = "herself" then
        word = "she"
      elseif word = "its" or word = "itself" then
        word = "it"
      elseif word = "whom" or word = "whose" then
        word = "who"
      elseif word = "one's" or word = "oneself" then
        word = "one"
      elseif word = "us" or word = "our" or word = "ours" or word = "ourselves" then
        word = "we"
      elseif word = "yours" or word = "your" or word = "yourself" or word = "yourselves" then
        word = "you"
      elseif word = "them" or word = "their" or word = "theirs" or word = "themselves" then
        word = "they"
      elseif word = "am" or word = "is" or word = "are" or word = "was" or word = "were" or word = "being" or word = "been" then
        word = "be"
      elseif word = "has" or word = "had" or word = "having" then
        word = "have"
      end if
      if dict.Exists(word) then
        dict(word) = dict(word) + 1
      else 
        dict.Add word, 1
      end if
    next
loop

wscript.echo "CHECKING THE ZIPF's LAW"
wscript.echo
wscript.echo "The first column is the number of corresponding words in the text and the second column is the number of words which should occur in the text according to the Zipf's law."
wscript.echo
wscript.echo "The most popular words in " + book + " are:"
wscript.echo

for i = 0 To amount
  key = "" : count = 0
  for each word in dict
    if dict(word) > count then
      key = word 
      count = dict(word)
    end if
  next
  if count <> 0 then 
    dict.remove(key)
    if key <> vbLf and key <> vbCr and key <> "" and key <> " " and i = 1 then
      zipfCount = count
      wscript.echo key, vbTab, count, vbTab, round(zipfCount / i)
    elseif key <> vbLf and key <> vbCr and key <> "" and key <> " " then
      wscript.echo key, vbTab, count, vbTab, round(zipfCount / i)
    end if
  end if
next

wscript.echo
wscript.echo "The most popular still remaining short forms in " + book + " are:"
wscript.echo

for i = 0 To amount
  key = "" : count = 0
  for each word in dict
    if dict(word) > count and InStr(word, "'") then
      key = word 
      count = dict(word)
    end if
  next
  if count <> 0 then 
    dict.remove(key)
    if key <> vbLf and key <> vbCr and key <> "" and key <> " " then
      wscript.echo key, vbTab, count
    end if
  end if
next