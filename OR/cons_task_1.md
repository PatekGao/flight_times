input:静态时刻表：每小时为一格。
包含数据：纵轴：0-23点；横轴：ARR_DOM,ARR_INT,DEP_DOM,DEP_INT
另有ARR_MAX,DEP_MAX,TOT_MAX, ARR_DOM_MAX,ARR_INT_MAX,DEP_DOM_MAX,DEP_INT_MAX, DOM_MAX,INT_MAX的动态小时最大值
ARR_LIMIT,DEP_LIMIT,TOT_LIMIT的动态每五分钟最大值

操作：将静态时刻表转换为动态时刻表：分裂为每五分钟为一格。
包含数据：横轴：0:00,0:05,0:10...以五分钟为间隔；横轴：ARR_DOM,ARR_INT,DEP_DOM,DEP_INT


限制条件1：
细分类别滑动小时加和限制（动态小时最大值）
滑动小时加和：
0:00-0:55：g1+g2+...+g12=m1
0:05-1:00：g2+g3+...+g13=m2
0:10-1:05：g3+g4+...+g14=m3
...
23:00-23:55：g277+...+g288=m277

m1...m277都需小于等于ARR_DOM_MAX
ARR_DOM滑动小时加和需要小于等于ARR_DOM_MAX
以此类推，ARR_DOM,ARR_INT,DEP_DOM,DEP_INT都需满足此条件。

限制条件2：
整合类别滑动小时加和限制（动态小时最大值）
ARR_DOM+ARR_INT=ARR
DEP_DOM+DEP_INT=DEP
ARR_DOM+DEP_DOM=DOM
ARR_INT+DEP_INT=INT
ARR_DOM+ARR_INT+DEP_DOM+DEP_INT=TOT
ARR滑动小时加和需要小于等于ARR_MAX
以此类推，ARR，DEP，DOM，INT，TOT

限制条件3：
动态小时最大值等于限制
ARR,DEP,TOT, ARR_DOM,ARR_INT,DEP_DOM,DEP_INT, DOM,INT
以ARR为例，在动态小时钟，至少有一个动态小时的ARR需要等于ARR_MAX
以此类推，ARR_BIG,DEP_BIG,TOT_BIG, ARR_DOM_BIG,ARR_INT_BIG,DEP_DOM_BIG,DEP_INT_BIG, DOM,INT_BIG
都需要等于相应的MAX

限制条件4：
动态每五分钟最大值限制
动态每五分钟最大值都需小于等于MIN_LIMIT
其中有ARR_LIMIT,DEP_LIMIT,TOT_LIMIT
每一个五分钟（即每一行）的ARR，DEP，TOT都需要满足小于等于相应的LIMIT
