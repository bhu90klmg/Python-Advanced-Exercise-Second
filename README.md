# Python-Advanced-Exercise-Second
將5個.csv檔案的cost值讀為一個向量，此向量的所有元素比較出最低者的索引值Index_low及其cost值cost_low;

最高者的索引值Index_high及其cost值cost_high，並印出。

然後根據以下數學方程式計算並變更part 1的行號.xlsx內容:

x_high = X23(:,Index_high)

x_low  = X23(:,Index_low)    

cost_high = cost(Index_high)

cost_low = cost(Index_low)

X23各列加總得S23 ; rand是0~1的隨機實數 ; "/4"是各列除上4 ; "*alfa"是各列乘上alfa ; 

"X23(:,1)"指第1行的向量內容被取代, "X23(:,2)"指第2行的向量內容被取代, ..., "X23(:,5)"指第5行的向量內容被取代:

x_base = (S23-x_high)/4

alfa = 0.9+rand*0.2

X23(:,1) = x_base+(x_base-x_high)*alfa

alfa = 0.9+rand*0.2

X23(:,2) = x_base+(x_base-x_high)*alfa

alfa = 0.9+rand*0.2

X23(:,3) = x_base+(x_base-x_high)*alfa

alfa = 0.9+rand*0.2

X23(:,4) = x_base+(x_base-x_high)*alfa

alfa = 0.9+rand*0.2

X23(:,5) = x_base+(x_base-x_high)*alfa

若該列超出該列所訂的數值範圍(part 1所示)，則以範圍的邊界值取代之，

然後用新的X23變更part 1的行號.xlsx內容，並開始"逐一變更"5個.csv檔案的當下time值及隨機cost值，

且新的每個.csv內的當下time值是不同的(因為是"逐一變更"的關係)。



讀入work0到work4的xlsx檔

都加入x23的陣列中

分別讀入time.csv檔，並且取出cost值存入向量

再cost的向量中找出最大與最小值的索引值

再根據索引值比對x23的x_high與x_low

根據題目的進行運算

利用openpyxl對原本的xlsx檔進行更改

最後利用pandas對time的時間進行時間的更改
