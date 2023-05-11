#!/usr/bin/env python
# coding: utf-8

# In[ ]:


Блок SQL
Задание 1
Условие
У вас есть две таблицы: "users" (с полями id, name, email) и "orders" (с полями id, user_id, order_date, total_price). 
Напишите запрос, который вернет (уникальные) имена пользователей, которые сделали заказы в течение последней недели.

Решение
-- заранее фильтруем таблицу и оставляем только тех пользователей, которые делали заказ в течение прошлой недели
with orders_last_week AS 
(
    select
         user_id
    from orders
    where    date_part('week',NOW()) - date_part('week',order_date) = 1 
             and date_part('year', order_date) = 2023
)

select distinct(u.name) 
from users u
join orders_last_week olw
on u.id = olw.user_id



Задание 2

Условие
У вас есть три таблицы: customers (id, name), orders (id, customer_id, order_date), и order_items (id, order_id, product_name, price, quantity). 
Напишите SQL-запрос, который выводит список клиентов (id, name) и общую сумму их покупок (total_spent), только для тех клиентов, которые сделали более одного заказа.

Решение
-- пользователи, которые сделали больше одного заказа
with filtered_orders AS (
    select 
    id,
           customer_id
    from (
        select 
               id,
               customer_id,
               count(id) over(partition by customer id) cnt_orders
        from orders
    )
    where cnt_orders > 1 
),

-- считаем сумму покупок в разрезе каждого заказа для пользователей, которые делали более одного заказа
total_spent_per_order AS (
    select 
           fo.customer_id,
           oi.order_id,
           sum(oi.price * oi.quantity) total_spent
    from (
         select *
         FROM order_items oi
         join filtered_orders fo
         on oi.id = fo.order_id

    )
    group by fo.customer_id,
             oi.order_id

),

-- финальный запрос, в котором мы считаем сумму трат, которая приходится на каждого пользователя и джойним его имя

select a.id,
       c.name,
       a.total_spent

from (
    select customer_id,
           sum(total_spent) total_spent
    from total_spent_per_order
    group by customer_id
) a

join customers c
on a.customer_id = c.id





Задание 3
Условие

В таблице "orders" хранятся данные о заказах в интернет-магазине. 
Необходимо написать запрос, который вернет количество заказов по месяцам за последние 6 месяцев, начиная с текущего месяца. 
Необходимо учитывать только те заказы, которые были оплачены.

Решение

Таблица "orders":
• id (int) - уникальный идентификатор заказа
• date (date) - дата заказа
• status (varchar) - статус заказа (оплачен, не оплачен и т.д.)
• amount (float) - сумма заказа
• customer_id (int) - идентификатор покупателя

select date_part('month',date) order_month,
       count(id) cnt_orders
from orders 
where (date_part('year', now()) - date_part('year',date)) * 12 + 
       (date_part('month', now()) - date_part('month', date)) <=6
       and status = 'paid'
group by date_part('month',date) order_month





Задание 4
Условие
У вас есть две таблицы: "orders" и "customers". В таблице "orders" содержится информация о заказах, включая идентификатор заказа, 
дату заказа, идентификатор клиента и общую сумму заказа. В таблице "customers" содержится информация о клиентах, 
включая идентификатор клиента, имя, фамилию, электронную почту и страну. Необходимо составить запрос, 
который выводит идентификаторы клиентов, их имя и фамилию, количество сделанных ими заказов и общую сумму этих заказов. 
Выведите только тех клиентов, которые сделали более 3 заказов и общая сумма заказов которых превышает 1000 долларов.

Поля таблицы "orders":
• order_id INT
• order_date DATE
• customer_id INT
• order_amount FLOAT

Поля таблицы "customers":
• customer_id INT
• first_name VARCHAR(50)
• last_name VARCHAR(50)
• email VARCHAR(50)
• country VARCHAR(50)




-- отфильтровываем таблицу с заказами и оставляем только id клиентов, которые отвечают условиям задачи

with filtered_customers as (
    select customer_id,
           sum(order_amount) total_sum,
           count(order_id) cnt_orders
    from orders
    group by customer_id
    having sum(order_amount) > 1000 and count(order_id) > 3

)

Решение

select c.first_name,
       c.last_name
       fc.total_sum,
       fc.cnt_orders
from customers c
join filtered_customers fc
on c.customer_id = fc.customer_id



Блок Python
Задание 1
Условие

Описание
У вас есть CSV-файл sales.csv со следующими полями:
• order_id - идентификатор заказа
• customer_id - идентификатор покупателя
• order_date - дата заказа (в формате YYYY-MM-DD)
• product_id - идентификатор продукта
• quantity - количество продуктов в заказе
• price - цена продукта
• discount - скидка на продукт (в процентах)

Ваша задача - написать скрипт на Python, который считывает данные из файла sales.csv, 
группирует их по дате заказа и продукту, и выводит таблицу, содержащую следующие столбцы:
• order_date - дата заказа
• product_id - идентификатор продукта
• total_quantity - общее количество продуктов в заказах
• total_sales - общая выручка с продаж (учитывая скидки)

Таблица должна быть отсортирована по дате заказа и продукту.

Решение

df = pd.read_csv('sales.csv') #что-то про путь

# предположим, что скидка хранится в формате 30%, тогда нам нужно перевести число в формат float, разделив на 100
# создаем столбец с ценой, учитывающей скидку

df['price_with_discount'] = df['price'] * df['discount'].div(100)

res = df.groupby(['order_date','product_id'],as_index=False)        .agg({'quntity':'sum','price_with_discount':'sum'})        .rename(columns={'quantity':'total_quantity','total_sales':'sales'})        .sort_values(by=['order_date','product_id'])


Задание 2
Условие

У вас есть два датафрейма, содержащих информацию о продажах товаров в разных магазинах. 
Первый датафрейм df_sales содержит следующие столбцы:

order_id - уникальный идентификатор заказа
product_id - уникальный идентификатор товара
store_id - уникальный идентификатор магазина
date - дата продажи товара в формате 'YYYY-MM-DD'
quantity - количество проданных товаров
price - цена за единицу товара
total_price - общая стоимость товара (количество * цена)

Второй датафрейм df_stores содержит следующие столбцы:

store_id - уникальный идентификатор магазина
city - город, в котором расположен магазин
region - регион, в котором расположен магазин
sales_rep - имя продавца в магазине

Вам необходимо сделать следующее:

-Загрузить данные из файлов sales.csv и stores.csv в датафреймы df_sales и df_stores соответственно.
-Найти общее количество проданных товаров, общее количество заказов и общую выручку.
-Рассчитать среднюю стоимость заказа.
-Рассчитать общую выручку и количество проданных товаров для каждого магазина.
-Найти общую выручку и количество проданных товаров для каждого региона.
-Найти топ-3 магазинов по выручке и количество проданных товаров.
-Найти топ-3 продавцов по выручке.


Решение

-- загружаем данные в отдельные датафреймы. Так как известны только названия датафреймов, полный путь до файлов не указываю.

df_sales = pd.read_csv('sales.csv')
df_stores = pd.read_csv('stores.csv')

-- общее количество проданных товаров, общее количество заказов, общая выручка

cnt_products = df_sales['quntity'].sum()
cnt_orders = df_sales['order_id'].count()
revenue = df_sales['total_price'].sum()

-Рассчитать среднюю стоимость заказа.

df_sales.groupby('order_id', as_index=False).agg({'total_price':'mean'})['total_price'].mean()

-Рассчитать общую выручку и количество проданных товаров для каждого магазина.

sales_by_store = df_sales.groupby('store_id').agg({'total_price':'sum', 'quantity':'sum'})

-Найти общую выручку и количество проданных товаров для каждого региона

df_regions = df_sales.merge(df_stores['region','store_id','sales_rep'], how='left', on ='store_id')
df_regions.groupby('region').agg({'total_price':'sum', 'quantity':'sum'})

-Найти топ-3 магазинов по выручке и количество проданных товаров

top_sales = sales_by_store.sort_values(by='total_price', ascending=False)[:3]
top_cnt = sales_by_store.sort_values(by='quantity', ascending=False)[:3]

-Найти топ-3 продавцов по выручке

sales_by_managers = df_regions.groupby('sales_rep').agg('total_price':'sum').sort_values(by='total_price',ascending=False)[:3]











