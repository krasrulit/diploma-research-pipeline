-- Parameter:
--   $date = first day of the reporting month, e.g. date '2026-03-01'
--
-- Output tables:
--   usr_wrk.region_active_snapshot_base
--   usr_wrk.region_active_segments
--   usr_wrk.region_active_product_wide
--   usr_wrk.region_active_products_long
--   usr_wrk.region_active_ent_products_long

drop table if exists rosbank;
create table rosbank as
with prepare as (
    select
        party_rk
    from usr_dep.dc_channel
    where subchannel_desc = 'Росбанк'
      and product_nm = 'Tinkoff Black'
),
prepare2 as (
    select
        ros.party_rk,
        ch.subchannel_desc,
        ch.utilization_dt,
        row_number() over (
            partition by ros.party_rk
            order by ch.utilization_dt, ch.crm_income_dttm
        ) as rn
    from prepare ros
    inner join usr_dep.dc_channel ch
        on ros.party_rk = ch.party_rk
       and ch.utilization_dt is not null
)
select
    party_rk
from prepare2
where rn = 1
  and utilization_dt >= date '2025-01-01'
  and utilization_dt < date '2025-02-01'
  and subchannel_desc = 'Росбанк'
distributed by (party_rk);

drop table if exists usr_wrk.region_active_snapshot_base;
create table usr_wrk.region_active_snapshot_base as
with params as (
    select
        $date::date as month_start,
        (date_trunc('month', $date::date) + interval '1 month' - interval '1 day')::date as month_end
),
raw as (
    select
        p.month_start as dt,
        p.month_end,
        y.party_rk,
        rm.bucket,
        rm.russian_region_nm,
        y.fst_product_utilization_dt,
        y.active_flg_stat_rep,
        y.pro_flg,
        y.premium_flg,
        y.active_flg_dk,
        y.active_flg_kk,
        y.active_flg_mvno,
        y.active_flg_invest,
        y.active_flg_investcop,
        y.active_flg_ns,
        y.active_flg_moneypot,
        y.active_flg_dep,
        y.active_flg_kn,
        y.active_flg_knz,
        y.active_flg_dolyami,
        y.active_flg_pos,
        y.active_flg_auto,
        y.active_flg_osago,
        y.active_flg_kasko,
        y.active_flg_avia,
        y.active_flg_hotels,
        y.active_flg_nfmcg,
        y.active_flg_grocery,
        y.active_flg_afisha,
        y.active_flg_fuel,
        y.active_flg_vzr,
        y.active_flg_sme,
        y.active_flg_junior,
        y.active_flg_sln,
        y.mau_flg,
        row_number() over (
            partition by p.month_end, y.party_rk
            order by rm.valid_from_dt desc nulls last, rm.bucket nulls last, rm.russian_region_nm nulls last
        ) as rn
    from params p
    inner join usr_acq.ykash_every_mnth_active4 y
        on y.somonth_dt = p.month_start
    inner join prod_v_emart.person_party pp
        on pp.party_rk = y.party_rk
       and pp.age >= 18
    left join rosbank r
        on r.party_rk = y.party_rk
    left join usr_dep.mami_party_x rm
        on rm.party_rk = y.party_rk
       and p.month_end >= rm.valid_from_dt
       and p.month_end <= rm.valid_to_dt
    where y.fst_product_utilization_dt is not null
      and y.active_flg_stat_rep = 1
      and (
            r.party_rk is null
         or (
                r.party_rk is not null
            and (
                    y.fst_product_utilization_dt < date '2025-01-01'
                 or y.fst_product_utilization_dt >= date '2025-02-01'
                )
            )
      )
)
select
    dt,
    month_end,
    party_rk,
    bucket,
    russian_region_nm,
    fst_product_utilization_dt,
    active_flg_stat_rep,
    pro_flg,
    premium_flg,
    active_flg_dk,
    active_flg_kk,
    active_flg_mvno,
    active_flg_invest,
    active_flg_investcop,
    active_flg_ns,
    active_flg_moneypot,
    active_flg_dep,
    active_flg_kn,
    active_flg_knz,
    active_flg_dolyami,
    active_flg_pos,
    active_flg_auto,
    active_flg_osago,
    active_flg_kasko,
    active_flg_avia,
    active_flg_hotels,
    active_flg_nfmcg,
    active_flg_grocery,
    active_flg_afisha,
    active_flg_fuel,
    active_flg_vzr,
    active_flg_sme,
    active_flg_junior,
    active_flg_sln,
    mau_flg
from raw
where rn = 1
distributed by (party_rk);

drop table if exists usr_wrk.region_active_segments;
create table usr_wrk.region_active_segments as
with params as (
    select
        $date::date as month_start,
        (date_trunc('month', $date::date) + interval '1 month' - interval '1 day')::date as month_end
),
enterp_snapshot as (
    select distinct
        s.customer_rk as party_rk
    from params p
    inner join prod_v_sse.smexl_base_sample_vers s
        on p.month_end between s.valid_from_dttm and s.valid_to_dttm
    where s.not_ip_ooo_flg = 0
      and s.not_client_flg = 0
)
select
    b.dt,
    b.month_end,
    b.party_rk,
    b.bucket,
    b.russian_region_nm,
    'Все активные клиенты' as segment
from usr_wrk.region_active_snapshot_base b

union all

select
    b.dt,
    b.month_end,
    b.party_rk,
    b.bucket,
    b.russian_region_nm,
    'Активные предприниматели' as segment
from usr_wrk.region_active_snapshot_base b
inner join enterp_snapshot e
    on e.party_rk = b.party_rk

union all

select
    b.dt,
    b.month_end,
    b.party_rk,
    b.bucket,
    b.russian_region_nm,
    'Активный антисегмент' as segment
from usr_wrk.region_active_snapshot_base b
left join enterp_snapshot e
    on e.party_rk = b.party_rk
where e.party_rk is null
distributed by (party_rk);

drop table if exists usr_wrk.region_active_product_wide;
create table usr_wrk.region_active_product_wide as
select
    s.dt,
    s.month_end,
    s.bucket,
    s.russian_region_nm,
    s.segment,
    count(distinct b.party_rk) as active_cnt,
    sum(b.pro_flg) as pro_flg,
    sum(b.premium_flg) as premium_flg,
    sum(b.active_flg_dk) as active_flg_dk,
    sum(b.active_flg_kk) as active_flg_kk,
    sum(b.active_flg_mvno) as active_flg_mvno,
    sum(b.active_flg_invest) as active_flg_invest,
    sum(b.active_flg_investcop) as active_flg_investcop,
    sum(b.active_flg_ns) as active_flg_ns,
    sum(b.active_flg_moneypot) as active_flg_moneypot,
    sum(b.active_flg_dep) as active_flg_dep,
    sum(b.active_flg_kn) as active_flg_kn,
    sum(b.active_flg_knz) as active_flg_knz,
    sum(b.active_flg_dolyami) as active_flg_dolyami,
    sum(b.active_flg_pos) as active_flg_pos,
    sum(b.active_flg_auto) as active_flg_auto,
    sum(b.active_flg_osago) as active_flg_osago,
    sum(b.active_flg_kasko) as active_flg_kasko,
    sum(b.active_flg_avia) as active_flg_avia,
    sum(b.active_flg_hotels) as active_flg_hotels,
    sum(b.active_flg_nfmcg) as active_flg_nfmcg,
    sum(b.active_flg_grocery) as active_flg_grocery,
    sum(b.active_flg_afisha) as active_flg_afisha,
    sum(b.active_flg_fuel) as active_flg_fuel,
    sum(b.active_flg_vzr) as active_flg_vzr,
    sum(b.active_flg_sme) as active_flg_sme,
    sum(b.active_flg_junior) as active_flg_junior,
    sum(b.active_flg_sln) as active_flg_sln,
    sum(b.mau_flg) as mau_flg
from usr_wrk.region_active_segments s
inner join usr_wrk.region_active_snapshot_base b
    on s.party_rk = b.party_rk
   and s.dt = b.dt
group by 1,2,3,4,5
distributed by (dt, segment);

drop table if exists usr_wrk.region_active_products_long;
create table usr_wrk.region_active_products_long as
with tmp as (
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт' as metric_type, 'Pro' as metric_nm, pro_flg as metric_value, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Premium', premium_flg, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'ДК', active_flg_dk, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'КК', active_flg_kk, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'MVNO', active_flg_mvno, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Инвестиции', active_flg_invest, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Инвесткопилка', active_flg_investcop, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Накопительный счет', active_flg_ns, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Копилка', active_flg_moneypot, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Депозит', active_flg_dep, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Кредит наличными', active_flg_kn, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Кредит наличными залоговый', active_flg_knz, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Долями', active_flg_dolyami, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'POS', active_flg_pos, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Авто', active_flg_auto, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'ОСАГО', active_flg_osago, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'КАСКО', active_flg_kasko, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Авиа', active_flg_avia, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Отели', active_flg_hotels, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'NFMCG', active_flg_nfmcg, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Grocery', active_flg_grocery, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Афиша', active_flg_afisha, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Топливо', active_flg_fuel, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'ВЗР', active_flg_vzr, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'SME', active_flg_sme, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'Junior', active_flg_junior, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'SLN', active_flg_sln, active_cnt from usr_wrk.region_active_product_wide
    union all
    select dt, month_end, bucket, russian_region_nm, segment, 'Продукт', 'MAU', mau_flg, active_cnt from usr_wrk.region_active_product_wide
),
portfolio as (
    select
        dt,
        month_end,
        segment,
        metric_type,
        metric_nm,
        sum(metric_value) as total_metric_value,
        sum(active_cnt) as total_active_cnt
    from tmp
    group by 1,2,3,4,5
)
select
    t.dt,
    t.month_end,
    t.bucket,
    t.russian_region_nm,
    t.segment,
    t.metric_type,
    t.metric_nm,
    t.metric_value,
    t.active_cnt,
    round(t.metric_value::numeric / nullif(t.active_cnt, 0), 6) as penetration,
    round(p.total_metric_value::numeric / nullif(p.total_active_cnt, 0), 6) as portfolio_penetration,
    round(((t.metric_value::numeric / nullif(t.active_cnt, 0)) / nullif(p.total_metric_value::numeric / nullif(p.total_active_cnt, 0), 0) - 1), 4) as deviation_from_portfolio,
    round((((t.metric_value::numeric / nullif(t.active_cnt, 0)) / nullif(p.total_metric_value::numeric / nullif(p.total_active_cnt, 0), 0) - 1) * 100), 2) as deviation_from_portfolio_pct,
    p.total_metric_value,
    p.total_active_cnt
from tmp t
left join portfolio p
    on t.dt = p.dt
   and t.month_end = p.month_end
   and t.segment = p.segment
   and t.metric_type = p.metric_type
   and t.metric_nm = p.metric_nm
distributed by (dt, segment);

drop table if exists usr_wrk.region_active_ent_products_long;
create table usr_wrk.region_active_ent_products_long as
select
    *
from usr_wrk.region_active_products_long
where segment = 'Активные предприниматели'
distributed by (dt, metric_nm);
