SELECT t1.代码, t1.名称, t1.yzlx,
CASE
WHEN t1.指标1 <> t1.指标2 + t1.指标3 + t1.指标5 THEN '期末存栏<>仔猪+待育肥猪+种猪'
WHEN t1.指标3 < t1.指标4 THEN '待育肥猪应大于等于50公斤以上'
WHEN t1.指标5 < t1.指标6 THEN '种猪应大于等于能繁母猪'
WHEN t1.指标7 < t1.指标8 + t1.指标9 THEN '期内增加应大于等于自繁加购进'
WHEN t1.指标10 <> t1.指标11 + t1.指标12 + t1.指标15 THEN '期内减少应等于自宰加出售头数加其他原因减少'
WHEN t1.指标15 < t1.指标16 THEN '其他原因减少应大于等于出售仔猪数量'
WHEN (t1.指标14 / NULLIF(t1.指标12, 0) > 200 OR t1.指标14 / NULLIF(t1.指标12, 0) < 70) AND t1.指标12 <> 0 THEN '检查肥猪重量是否在合理范围'
WHEN (t1.指标13 / NULLIF(t1.指标14, 0) > 20 OR t1.指标13 / NULLIF(t1.指标14, 0) < 8) AND t1.指标14 <> 0 THEN '检查出售单价是否在合理范围'
WHEN t1.指标1 <> t2.指标1 + t1.指标7 - t1.指标10 THEN '期末存栏不等于上期存栏加本期增加减本期减少'
WHEN (t1.指标12 + t1.指标11 > t2.指标4) THEN '本期自宰加出栏不应大于上期50公斤以上待育肥猪'
END AS 错误类型
FROM 畜禽调查表 t1
LEFT JOIN 畜禽调查表 t2 ON t1.代码 = t2.代码
AND t1.yf = '04' AND t1.yzlx = N'1'
AND t2.yf = '03' AND t2.yzlx = N'1'
WHERE
yzlx = '1'
AND yf = N'04'