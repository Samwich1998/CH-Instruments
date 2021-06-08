function [baseline] = getbaseline(data,order,iterations)
l=length(data);
x1 = 1:l;
for i = 1 : iterations
    p = polyfit(x1,data,order);
    baseline = polyval(p,x1);
    for j = 1 :  l

       if data(j)>baseline(j)
           data(j)=baseline(j);
       else
           data(j)=data(j);
       end
    end
end       
end

