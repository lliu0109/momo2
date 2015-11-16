function Y1 = interpolation_1d(X,Y,X1)
%this function does the linear interpolation in one dimention 
%make sure X and Y are monotonic increasing

if length(X) ~= length(Y) 
     error('X and Y length must agree');
end

Y1 = zeros(size(X1));


for i = 1:length(X1)
    pos = find(X>X1(i), 1, 'first');
    if isempty(pos)
        Y1(i) = Y(end);
    else 
        if pos==1
            Y1(i) = Y(1);
        else
            if X1(i) == X(pos-1)
                Y1(i) = Y(pos-1);
            else
                Y1(i) = interp1(X,Y,X1(i));
            end
        end
    end
end
end


