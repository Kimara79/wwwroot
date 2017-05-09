<%
'1组数字排序
function orders1(arr)
 dim i,j,tmp
   for i=0 to UBound(arr)
     for j=0 to UBound(arr)
       if int(arr(i))<int(arr(j)) then
          tmp=arr(i)
          arr(i)=arr(j)
          arr(j)=tmp
       end if
     next
   next
 orders1=arr
end function

'1组字符排序
function orders1s(arr)
 dim i,j,tmp
   for i=0 to UBound(arr)
     for j=0 to UBound(arr)
       if arr(i)<arr(j) then
          tmp=arr(i)
          arr(i)=arr(j)
          arr(j)=tmp
       end if
     next
   next
 orders1s=arr
end function

'2组时间排序,按第1组排序
function orders2t(arr1,arr2)
 dim i,j,tmp
   for i=0 to UBound(arr1)-1
     for j=0 to UBound(arr1)-1
       if DateDiff("s",arr1(i),arr1(j))>0  then
          tmp=arr1(i)
          arr1(i)=arr1(j)
          arr1(j)=tmp
		  
          tmp=arr2(i)
          arr2(i)=arr2(j)
          arr2(j)=tmp
       end if
     next
   next
   for i=0 to UBound(arr1)-1
    orders2t=orders2t&arr1(i)&"|"&arr2(i)&";"
   next
end function

'2组字符排序,按第1组排序
function orders2s(arr1,arr2,arr3,arr4,arr5)
 dim i,j,tmp
   for i=0 to UBound(arr1)-1
     for j=0 to UBound(arr1)-1
       if arr1(i)<arr1(j) then
          tmp=arr1(i)
          arr1(i)=arr1(j)
          arr1(j)=tmp
		  
          tmp=arr2(i)
          arr2(i)=arr2(j)
          arr2(j)=tmp
       end if
     next
   next
   for i=0 to UBound(arr1)-1
    orders2s=orders5s&arr1(i)&"|"&arr2(i)&"|"&arr3(i)&"|"&arr4(i)&"|"&arr5(i)&";"
   next
end function
%>