function english = chinese2english(chinese) 

english = '';

for i = 1:length(chinese)
   t = double(chinese(i));
   if t == double(' ')
       continue;
   end
   if (t<=double('Z') && t>=double('A')) ||...
           (t<=double('z') && t>=double('a')) ...
           || t == double('_')
       english = [english chinese(i)];
   else
       english = [english '0x' dec2hex(t)];
   end
end

if isempty(english)
    english = 'x';
elseif double('0') <= double(english(1)) && ...
        double(english(1)) <= double('9')
    english = ['x' english];
end

