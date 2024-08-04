a = '2021-11-29 15:29:06'
b = '2021-11-29 15:35:26'

a_temp = a.split(" ")[1].split(":")
b_temp = b.split(" ")[1].split(":")

a_T = a_temp[0]
a_M = a_temp[1]
a_S = a_temp[2]

b_T = b_temp[0]
b_M = b_temp[1]
b_S = b_temp[2]

r_S = b_S - a_S
r_M = b_M - a_M
r_T = b_T - a_T

if(r_S < 0)
    {
        r_M -= 1
        r_S += 60
    }

if(r_M < 0)
    {
        r_T -= 1
        r_M += 60
    }

result = r_T + "시간" + r_M + "분" + r_S + "초"

console.log(result)