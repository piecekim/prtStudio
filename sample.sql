select 
	*
from 
	svr7254 s 
where 1=1
and S.종목코드 = 'A005930'
order by 
	일자
;

-- 1. 당일매수매도량 TodaySellAndBuy
select 
	개인
from
	svr7254 s 
where 1=1
and S.종목코드 = 'A005930'
and S.일자 = :입력일자
;

-- 3. 매집수량, 매집고점, 매집저점
select
	AA.매집수량 
	, AA.매집수량 
	, AA.매집수량 
from 
(
	select 
		A.*
	from
		추세_개인 A
	where 1=1
	and A.일자 < '20040101'
	and A.종목코드 = 'A005930'
	order by
		A.일자 desc
) AA
LIMIT 1
;

-- 4. XX일 추세
select
	AVG(AA.매집수량) as XX일추세
from 
(
	select 
		A.*
	from
		추세_개인 A
	where 1=1
	and A.일자 < '20100101'
	and A.종목코드 = 'A005930'
	order by
		A.일자 desc
) AA
LIMIT XX
;

select 
--		(ROW_NUMBER() OVER()) as rownum
	A.*
from
	svr7254 A
where 1=1
and A.일자 < '20100101'
and A.종목코드 = 'A005930'
order by
	A.일자 desc
;

select 
	종목코드
	, 일자
from
	svr7254 s 
where 1=1
and s.종목코드 = 'A000020'
group by
	일자, 종목코드
order by 
	일자
;

select 
	*
from
	종목코드
where 1=1
and 종목코드 = 'A000020'



Individual supply and demand trend analysis

SupplyAnaltsis_indi.py