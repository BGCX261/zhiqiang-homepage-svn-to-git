function quotedText = addQuotes(text)

if strcmp(class(text), 'double')
    text = num2str(text);
end

quotedText = [' ''', text, ''' '];