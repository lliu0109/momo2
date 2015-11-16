function Z1 = interpolation_2d(X,Y,Z,X1,Y1)
%this function does the linear interpolation in two dimentions 

[Z_row Z_col] = size(Z);
if Z_row ~= length(X) || Z_col ~= length(Y)
     error('X, Y and Z size must agree');
end

if length(X1) ~= length(Y1)
     error('Input X1, Y1 size must agree');
end

Z1 = zeros(size(X1));


for i = 1:length(X1)
    x_pos = find(X>X1(i), 1, 'first');
    y_pos = find(Y>Y1(i), 1, 'first');

    if isempty(x_pos)
        if isempty(y_pos)
            Z1(i) = Z(end, end);
        else
            if y_pos == 1
                Z1(i) = Z(end, 1);
            else
                %Z1 = interpolation_1d(Y, Z(end, :), Y1);
                Z1(i) = interp1(Y, Z(end,:), Y1(i));
            end
        end
    else
        if x_pos == 1
            if isempty(y_pos)
                Z1(i) = Z(1, end);
            else
                if y_pos == 1
                    Z1(i) = Z(1,1);
                else
                    %Z1 = interpolation_1d(Y, Z(1, :), Y1);
                    Z1(i) = interp1(Y, Z(1,:), Y1(i));
                end
            end
        else
            if isempty(y_pos)
                %Z1 = interpolation_1d(X, Z(:, end), X1);
                Z1(i) = interp1(X, Z(:,end), X1(i));
            else
                if y_pos == 1
                    %Z1 = interpolation_1d(X, Z(:, 1), X1);
                    Z1(i) = interp1(X, Z(:,1), X1(i));
                else
                    Z1(i) = interp2(Y,X,Z, Y1(i), X1(i));
                end
            end
        end
    end
end

end
