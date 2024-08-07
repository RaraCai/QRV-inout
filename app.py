import streamlit as st
import pandas as pd
from streamlit_echarts import st_pyecharts
import streamlit.components.v1 as components
from pyecharts import options as opts
from pyecharts.charts import Bar,Line,Pie,PictorialBar
from pyecharts.globals import SymbolType,ThemeType
from datetime import datetime,date,timedelta
from collections import Counter
import json

# 确保屏幕自适应
#import win32api
WIDTH=1400
HEIGHT=960


# ------------------------------------数据筛选----------------------------------------
# 依据机构和任意时间范围筛选数据入口
def data_filtering(date_key,filter_org=True,filter_date=True):
    # 删去测试数据
    df1=df[~df['任务名称'].str.contains('测试',na=False)]

    blank,box_org,box_date=st.columns([8,1,1])
    # 过滤日期
    if filter_date:
        with box_date:
            st.popover('时间范围').date_input(
                '请选择起止日期',[date(date.today().year,1,1),date.today()],key=date_key
        )
        date_range=st.session_state[date_key]
        if len(date_range)!=2:
            date_range=(date(date.today().year,1,1),date.today())
        # 筛选时间范围
        start=datetime(date_range[0].year,date_range[0].month,date_range[0].day,0,0,0)
        end=datetime(date_range[1].year,date_range[1].month,date_range[1].day,23,59,59)
        df2=df1[(df1['呼出开始时间']>=start) & (df1['呼出开始时间']<=end)]

    # 过滤机构
    if filter_org:    
        with box_org:
            st.popover('试点选择').selectbox(
                '试点选择',
                [
                    '全部试点',
                    '明楼', 
                    '南码头社区卫生服务中心',
                    '北蔡社区卫生服务中心'
                ],
                index=0,    # 不选默认全部试点
                key='organization'
            )
        organization=st.session_state['organization']
        if organization=='全部试点':
            filtered_df=df2
        else:
            filtered_df=df2[df2['机构名称']==organization]
    else:
        filtered_df=df2

    # 显示数据选择情况
    if filter_org:
        st.info(f'【当前选择试点】：{organization}')
    if filter_date:
        st.info(f'【当前选择时间范围】：{date_range[0]}~{date_range[1]}')
    
    return filtered_df
# 选择变化频率
def render_freq_selectbox(key):
    st.popover("时间频度").radio(
        '请选择数据变化的时间频度',
        [ "日度","周度","月度"],
        index=1,
        key=key
    )
    freq=st.session_state[key]
    
    if freq=='月度':
        freq='ME'
    elif freq=='周度':
        freq='W-MON'
    elif freq=='日度':
        freq='D'
    
    return freq


# -----------------------------------------------------------------------------------


# ------------------------------------数据查询处理------------------------------------
# 数据总览
def query_general(filtered_df:pd.DataFrame,freq):
    # 总数
    total=pd.DataFrame(
        {
            '呼出总数':[filtered_df.shape[0]],
            '通话总时长':[(filtered_df['通话时长'].sum()/60).round(2)],
            '呼出成功':[(filtered_df['呼出结果']=='呼出成功').sum()],
            '呼出失败':[(filtered_df['呼出结果']=='呼出失败').sum()],
            '执行任务总数':[filtered_df['任务名称'].nunique()],
        }
    )
    # 计算成功率、失败率
    total['成功率']=(total['呼出成功']/total['呼出总数']).fillna(0)
    total['成功率']=[f'{rate*100:.2f}%' for rate in total['成功率']]
    total['失败率']=(total['呼出失败']/total['呼出总数']).fillna(0)
    total['失败率']=[f'{rate*100:.2f}%' for rate in total['失败率']]
    # 计算重呼率
    task_counts=filtered_df['任务名称'].value_counts()
    repeat=task_counts[task_counts>1].index
    filtered_with_repeats = filtered_df[filtered_df['任务名称'].isin(repeat)]
    repeat_count = filtered_with_repeats['任务名称'].nunique()
    total['重呼率']=(repeat_count/total['执行任务总数']).fillna(0)
    total['重呼率']=[f'{rate*100:.2f}%' for rate in total['重呼率']]

    # 平均数-按freq
    if freq=='ME': interval=30
    elif freq=='W-MON':interval=7
    else: interval=1
    avg=filtered_df.groupby(pd.Grouper(key='呼出开始时间',freq=freq)).agg(
        平均呼出=('呼出结果','size'),
        通话平均时长=('通话时长','mean'),
        平均呼出成功=('呼出结果',lambda x:(x=='呼出成功').mean()),
        平均呼出失败=('呼出结果',lambda x:(x=='呼出失败').mean()),
        执行任务平均数=('任务类型',lambda x:x.nunique()/interval)
    )
    # 修正平均值
    avg['通话平均时长']=avg['通话平均时长'].fillna(0).round(2)
    avg['平均呼出成功']=avg['平均呼出成功'].fillna(0).round(2)
    avg['平均呼出失败']=avg['平均呼出失败'].fillna(0).round(2)
    avg['执行任务平均数']=avg['执行任务平均数'].fillna(0).round(2)

    # 修正索引
    if freq=='ME':
        avg.index=avg.index.strftime('%Y-%m')
    elif freq=='W-MON':
        tmp=pd.DataFrame()
        tmp['周开始日期']=avg.index
        tmp['周结束日期']=avg.index+pd.offsets.Week(weekday=6)
        tmp['周范围']=tmp['周开始日期'].astype('str')+'~'+(tmp['周结束日期']).astype('str')
        avg.index=tmp['周范围']
    elif freq=='D':
        avg.index=avg.index.strftime('%Y-%m-%d')
    avg.index.name='时间'

    return total,avg

# 呼出详情-人次数据
def query_detail_fig(filtered_df:pd.DataFrame,freq):
    stats=filtered_df.groupby(pd.Grouper(key='呼出开始时间',freq=freq)).agg(
        呼出总数=('呼出结果','size'),
        呼出成功=('呼出结果',lambda x:(x=='呼出成功').sum()),
        呼出失败=('呼出结果',lambda x:(x=='呼出失败').sum()),
        通话总时长=('通话时长','sum'),
        通话平均时长=('通话时长','mean')
    )
    # 修正成功率、失败率、呼出平均时长
    stats['成功率'] = (stats['呼出成功'] / stats['呼出总数']).fillna(0)
    stats['失败率'] = (stats['呼出失败'] / stats['呼出总数']).fillna(0) 
    stats['通话平均时长'] = stats['通话平均时长'].fillna(0) 

    stats['成功率']=[f'{rate*100:.2f}%' for rate in stats['成功率']]
    stats['失败率']=[f'{rate*100:.2f}%' for rate in stats['失败率']]
    
    # 时长类数据转换为分钟
    stats['通话总时长']=(stats['通话总时长']/60).round(2)
    stats['通话平均时长']=(stats['通话平均时长']/60).round(2)


    # 按freq重置索引
    if freq=='ME':
        stats.index=stats.index.strftime('%Y-%m')
    elif freq=='W-MON':
        tmp=pd.DataFrame()
        tmp['周开始日期']=stats.index
        tmp['周结束日期']=stats.index+pd.offsets.Week(weekday=6)
        tmp['周范围']=tmp['周开始日期'].astype('str')+'~'+(tmp['周结束日期']).astype('str')
        stats.index=tmp['周范围']
    elif freq=='D':
        stats.index=stats.index.strftime('%Y-%m-%d')
    stats.index.name='时间'

    return stats

# 呼出详情-任务类型数据
def query_task_fig(filtered_df:pd.DataFrame,freq):
    stats=filtered_df.groupby(pd.Grouper(key='任务类型')).agg(
        通话总时长=('通话时长','sum'),
        通话平均时长=('通话时长','mean'),
    )
    # 转换为分钟
    stats['通话总时长']=(stats['通话总时长']/60).round(2)
    stats['通话平均时长'] = stats['通话平均时长'].fillna(0) 
    stats['通话平均时长']=(stats['通话平均时长']/60).round(2)
    total=(stats['通话总时长'].sum())
    stats['占比']=(stats['通话总时长']/total) if total>0 else 0
    stats['占比']=[f'{rate*100:.2f}%' for rate in stats['占比']]
    
    return stats

# 成功失败详情
def query_recall_reason(filtered_df:pd.DataFrame,freq):
    # 成功
    def my_agg(group):
        result = pd.Series({  
            '一次成功': ((group['呼出结果'] == '呼出成功') & (group['重呼次数'] == 0)).sum(),  
            '一次重呼成功': ((group['呼出结果'] == '呼出成功') & (group['重呼次数'] == 1)).sum(),  
            '两次重呼成功': ((group['呼出结果'] == '呼出成功') & (group['重呼次数'] == 2)).sum()  
        })  
        return result
    
    group=filtered_df.groupby(pd.Grouper(key='呼出开始时间',freq=freq))
    suc=group.apply(my_agg)

    #修正索引
    if freq=='ME':
        suc.index=suc.index.strftime('%Y-%m')
    elif freq=='W-MON':
        tmp=pd.DataFrame()
        tmp['周开始日期']=suc.index
        tmp['周结束日期']=suc.index+pd.offsets.Week(weekday=6)
        tmp['周范围']=tmp['周开始日期'].astype('str')+'~'+(tmp['周结束日期']).astype('str')
        suc.index=tmp['周范围']
    elif freq=='D':
        suc.index=suc.index.strftime('%Y-%m-%d')
    suc.index.name='时间'

    # 失败
    fail=filtered_df.groupby(pd.Grouper(key='失败原因')).agg(
        人次=('失败原因','count')
    )
    fail['占比']=(fail['人次']/fail['人次'].sum()) if fail['人次'].sum()>0 else 0
    fail['占比']=[f'{rate*100:.2f}%' for rate in fail['占比']]

    return suc,fail


# 受种者回复数据
def query_reply_detail(data:pd.DataFrame,freq):
    reply=data.groupby(pd.Grouper(key='呼出开始时间',freq=freq)).agg(
        总回复=('按键回复','count'),
        第1次回复=('按键回复',lambda x: x.apply(lambda y: y.count('：') == 1 if isinstance(y, str) else False).sum()),
        第2次回复=('按键回复',lambda x: x.apply(lambda y: y.count('：') == 2 if isinstance(y, str) else False).sum()),
        第3次回复=('按键回复',lambda x: x.apply(lambda y: y.count('：') == 3 if isinstance(y, str) else False).sum())
    )
    reply['第1次占比']=(reply['第1次回复']/reply['总回复']).fillna(0)
    reply['第2次占比']=(reply['第2次回复']/reply['总回复']).fillna(0)
    reply['第3次占比']=(reply['第3次回复']/reply['总回复']).fillna(0)

    return reply



# -----------------------------------------------------------------------------------


# ------------------------------------前端图表绘制------------------------------------
# 数据总览-总数板块
def render_general_total(data:pd.DataFrame):
    if len(data)==0:
        return
    total=data.values.tolist()[0]
    col1,col2,col3,col4=st.columns(4)
    col1.metric('呼出总数',f'{total[0]}人次')
    col1.write(f'通话总时长 {total[1]}分钟')

    col2.metric('呼出成功',f'{total[2]}人次')
    col2.write(f'成功率 {total[5]}')

    col3.metric('呼出失败',f'{total[3]}人次')
    col3.write(f'失败率 {total[6]}')

    col4.metric('执行任务',f'{total[4]}个')
    col4.write(f'重呼率 {total[7]}')
    

# 数据总览-平均板块
def render_general_avg(data:pd.DataFrame):
    if len(data)==0:
        return
    # 数据处理
    length=len(data)
    if(length>0):
        avg=data.values.tolist()[0]
        if(length==1):
            pre=[0,0,0,0,0]
        else:
            pre=data.values.tolist()[1]
    else:
        avg=[0,0,0,0,0]
        pre=[0,0,0,0,0]
    
    # 渲染
    col1,col2,col3,col4=st.columns(4)
    col1.metric('平均呼出',f'{avg[0]}人次',f'{avg[0]-pre[0]:.2f}人次')
    if avg[1]>pre[1]:
        string=f'通话平均时长 {avg[1]}分钟  ↑{avg[1]-pre[1]:.2f}分钟'
    elif avg[1]<pre[1]:
        string=f'通话平均时长 {avg[1]}分钟  ↓{pre[1]-avg[1]:.2f}分钟'
    else:
        string=f'通话平均时长 {avg[1]}分钟  -'
    col1.write(string)

    col2.metric('平均呼出成功',f'{avg[2]}人次',f'{avg[2]-pre[2]:.2f}人次')
    col3.metric('平均呼出失败',f'{avg[3]}人次',f'{avg[3]-pre[3]:.2f}人次')
    col4.metric('平均执行任务',f'{avg[4]}个',f'{avg[4]-pre[4]:.2f}个')

# 呼出详情-人次板块
def render_calls_detail(data:pd.DataFrame):
    if len(data)==0:
        return
    show_suc_rate=[float(rate.strip('%')) for rate in data['成功率']]
    show_fail_rate=[float(rate.strip('%')) for rate in data['失败率']]

    # 柱形图
    bar=(
        Bar(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.8}px',height=f'{HEIGHT*0.4}px'))
        .add_xaxis(data.index.astype('str').tolist())
        .add_yaxis('呼出总数',data['呼出总数'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
        .add_yaxis('呼出成功',data['呼出成功'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
        .add_yaxis('呼出失败',data['呼出失败'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            tooltip_opts=opts.TooltipOpts(
                trigger='axis',
                axis_pointer_type='cross'
            )
        )
    )
    # 条形图
    line=(
        Line(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.8}px',height=f'{HEIGHT*0.4}px'))
        .add_xaxis(data.index.astype('str').tolist())
        .extend_axis(yaxis=opts.AxisOpts(type_='value',position='right',name='百分比'))
        .add_yaxis('成功率',show_suc_rate,label_opts=opts.LabelOpts(is_show=False),yaxis_index=1,color='green')
        .add_yaxis('失败率',show_fail_rate,label_opts=opts.LabelOpts(is_show=False),yaxis_index=1,color='red')
        .set_global_opts(
            tooltip_opts=opts.TooltipOpts(
                trigger='axis',
                axis_pointer_type='cross',
            ),
            datazoom_opts=opts.DataZoomOpts(),
            title_opts=opts.TitleOpts(title='通话人次变化曲线'),
            legend_opts=opts.LegendOpts(type_='scroll')
        )
    )
    # 组合绘图
    grid_html=line.overlap(bar)
    components.html(grid_html.render_embed(),width=WIDTH*0.8,height=HEIGHT*0.4)

# 呼出详情-时长板块
def render_duration_detail(data:pd.DataFrame):
    if len(data)==0:
        return
    bar=(
        Bar(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.3}px',height=f'{HEIGHT*0.3}px'))
        .add_xaxis(data.index.astype('str').tolist())
        .add_yaxis('通话平均时长',data['通话平均时长'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            tooltip_opts=opts.TooltipOpts(
                trigger='axis',
                axis_pointer_type='cross'
            )
        )
    )
    line=(
        Line(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.3}px',height=f'{HEIGHT*0.3}px'))
        .add_xaxis(data.index.astype('str').tolist())
        .add_yaxis('通话总时长',data['通话总时长'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
            tooltip_opts=opts.TooltipOpts(
                trigger='axis',
                axis_pointer_type='cross',
            ),
            datazoom_opts=opts.DataZoomOpts(),
            title_opts=opts.TitleOpts(title='总体'),
            legend_opts=opts.LegendOpts()
        )
    )
    # 组合绘图
    components.html(line.overlap(bar).render_embed(),width=WIDTH*0.3,height=HEIGHT*0.3)

# 呼出详情-按不同任务类型分类
def render_task_detail(task:pd.DataFrame):
    if len(data)==0:
        return
    bar=(
        Bar(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.3}px',height=f'{HEIGHT*0.3}px'))
        .add_xaxis(task.index.astype('str').tolist())
        .add_yaxis('通话总时长',task['通话总时长'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
        .reversal_axis()
        .set_global_opts(
            tooltip_opts=opts.TooltipOpts(
                trigger='axis',
                axis_pointer_type='cross',
            ),
            title_opts=opts.TitleOpts(title='按任务类型'),
            legend_opts=opts.LegendOpts()
        )
    )
    components.html(bar.render_embed(),width=WIDTH*0.3,height=HEIGHT*0.3)

# 月度数据-时段统计和疫苗分类
def render_monthly(data:pd.DataFrame):
    if len(data)==0:
        return
    # 月度的小时变化数据
    month_hour_df=[]
    # 月度疫苗呼出分类
    month_vac_df=[]
    # 月份列表
    month_list=[]
    for month,mdata in data.groupby(data['呼出开始时间'].dt.month):
        # 小时变化
        hourly=mdata.groupby(mdata['呼出开始时间'].dt.hour).agg(
            计数=('呼出结果','size'),
            成功数=('呼出结果',lambda x:(x=='呼出成功').sum())
        )
        hourly['呼出成功率']=(hourly['成功数']/hourly['计数']).fillna(0)
        hourly.index=hourly.index.astype('str')+'点'
        month_hour_df.append(hourly)
        
        # 疫苗分类
        month_vac_df.append(
            mdata.groupby(pd.Grouper(key='疫苗名称')).agg(数量=('疫苗名称','count'))
        )

        # 月份列表
        month_list.append(f'{month}月')

    # 选择月份
    st.popover("选择月份").radio(
        '选择月份',
        month_list,
        index=0,
        key='month'
    )
    month_val=int(st.session_state['month'].strip('月')) if st.session_state['month'] else 1
    hourly_df=month_hour_df[month_val-int(month_list[0].strip('月'))] if len(month_list) else pd.DataFrame()
    vaccine_df=month_vac_df[month_val-int(month_list[0].strip('月'))] if len(month_list) else pd.DataFrame()

    # 绘图
    with st.container(border=True):
        left,right=st.columns([1,1.5])
        with left:
            if(len(hourly_df)):
                style={
                '计数':'{0:.1f}人次',
                '成功数':'{0:.1f}人次',
                '呼出成功率':'{0:.2%}'
                }
                table=hourly_df.style.format(style).background_gradient(subset=['计数'],cmap='Greens').highlight_max(subset=['呼出成功率'],props='background-color:pink')
                st.write(f'{month_val}月呼出时段统计')
                st.caption('单击列名查看升序或降序结果')
                st.dataframe(table,width=int(WIDTH*0.3),height=int(HEIGHT*0.2))

        with right:
            if len(vaccine_df):
                pie=(
                    Pie(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.3}px',height=f'{HEIGHT*0.3}px'))
                        .add(
                            '',[list(z) for z in zip(vaccine_df.index.tolist(),vaccine_df['数量'].values.tolist())],
                            radius=['40%','75%'],
                        )
                        .set_global_opts(
                            legend_opts=opts.LegendOpts(type_='scroll',pos_top='bottom'),
                            title_opts=opts.TitleOpts(title=f'{month_val}月呼出疫苗分类',subtitle=f'总计：{vaccine_df.sum().values}人次'),
                        )
                    )
                components.html(pie.render_embed(),width=WIDTH*0.3,height=HEIGHT*0.3)
                


# 月度呼出人次与成功率
def render_monthly_suc(data:pd.DataFrame):
    if len(data)==0:
        return
    monthly=data.groupby(pd.Grouper(key='呼出开始时间',freq='ME')).agg(
        呼出人次=('呼出结果','size'),
        成功数=('呼出结果',lambda x:(x=='呼出成功').sum())
    )
    monthly['成功率']=(monthly['成功数']/monthly['呼出人次']).fillna(0)
    monthly['成功率']=[f'{rate*100:.2f}%' for rate in monthly['成功率']]
    
    monthly.index=monthly.index.strftime('%Y-%m')
    
    show_suc_rate=[float(rate.strip('%')) for rate in monthly['成功率']]
    picbar=PictorialBar(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.3}px',height=f'{HEIGHT*0.4}px'))
    picbar.add_xaxis(monthly.index.astype('str').tolist())
    picbar.add_yaxis(
            '呼出人次',
            monthly['呼出人次'].values.tolist(),
            label_opts=opts.LabelOpts(is_show=True,position='right'),
            symbol_size=15,symbol_repeat='fixed',is_symbol_clip=True,symbol=SymbolType.DIAMOND,
            symbol_offset=[0,10]
        )
    picbar.add_yaxis(
            '成功率',
            show_suc_rate,
            label_opts=opts.LabelOpts(is_show=False),
            symbol_size=15,symbol_repeat='fixed',is_symbol_clip=True,symbol=SymbolType.ROUND_RECT,
            symbol_offset=[0,-10],
        )
    picbar.reversal_axis()
    picbar.set_global_opts(
            title_opts=opts.TitleOpts(title='月呼出人次与成功率'),
            xaxis_opts=opts.AxisOpts(is_show=False),
            yaxis_opts=opts.AxisOpts(
                axistick_opts=opts.AxisTickOpts(is_show=False),
                axisline_opts=opts.AxisLineOpts(linestyle_opts=opts.LineStyleOpts(opacity=0))
            ),
            legend_opts=opts.LegendOpts(pos_top='bottom')
        )
    
    components.html(picbar.render_embed(),width=WIDTH*0.3,height=HEIGHT*0.4)
    




# 失败详情-失败原因分类
def render_fail_detail(data:pd.DataFrame):
    if len(data)==0:
        return
    pie=(
        Pie(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.38}px',height=f'{HEIGHT*0.35}px'))
        .add(
            '',[list(z) for z in zip(data.index.tolist(),data['人次'].values.tolist())],
            radius=['40%','75%'],
            label_opts=opts.LabelOpts(is_show=True)
        )
        .set_global_opts(
            legend_opts=opts.LegendOpts(type_='scroll',pos_top='bottom'),
            title_opts=opts.TitleOpts(title='失败原因')
        )
    )

    components.html(pie.render_embed(),width=WIDTH*0.4,height=HEIGHT*0.38)

# 成功详情-重呼详情
def render_suc_detail(data:pd.DataFrame):
    
    if len(data)==0:
        return

    bar=Bar(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.38}px',height=f'{HEIGHT*0.4}px')).add_xaxis(data.index.tolist())
    bar.width='680px'
    bar.height='360px'
    bar.add_yaxis('一次成功',data['一次成功'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
    bar.add_yaxis('一次重呼成功',data['一次重呼成功'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
    bar.add_yaxis('两次重呼成功',data['两次重呼成功'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
    bar.reversal_axis()
    bar.set_global_opts(
        tooltip_opts=opts.TooltipOpts(
        trigger='axis',
        axis_pointer_type='cross',
        ),
        datazoom_opts=opts.DataZoomOpts(yaxis_index=0,orient='vertical',pos_left='right',type_='slider'),
        title_opts=opts.TitleOpts(title='重呼详情'),
        legend_opts=opts.LegendOpts(pos_bottom='bottom'),
        yaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=20))
    )
    
    components.html(bar.render_embed(),width=WIDTH*0.38,height=HEIGHT*0.4)

# 回复详情=按键回复内容分类
def render_reply_classify(data:pd.DataFrame):
    if len(data)==0:
        return
    # 提取和清理数据
    texts=[]

    for cell in data['按键回复']:
        if isinstance(cell, str):
            # 提取最后一个 `：` 后的文本
            last_colon_text = cell.split('：')[-1].split('；')[0].strip()
            if last_colon_text and not last_colon_text.isdigit() and last_colon_text != '#':
                texts.append(last_colon_text)

            # 提取 `：` 和 `；` 之间的文本
            parts = cell.split('：')
            if len(parts) > 1:
                middle_texts = [part.strip() for part in parts[1].split('；') if
                                part.strip() and not part.strip().isdigit() and part.strip() != '#']
                texts.extend(middle_texts)
        
    # 统计文本频率
    text_counts = Counter(texts)
    # 绘图
    pie=(
        Pie(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.38}px',height=f'{HEIGHT*0.35}px'))
        .add(
            '',[list(z) for z in zip(text_counts.keys(),text_counts.values())],
            radius=['40%','75%'],
            label_opts=opts.LabelOpts(is_show=True)
        )
        .set_global_opts(
            legend_opts=opts.LegendOpts(is_show=False),
            title_opts=opts.TitleOpts(title='回复内容分类'),
        )
    )
    components.html(pie.render_embed(),width=WIDTH*0.4,height=HEIGHT*0.38)
    with st.expander('详细数据'):
        details=pd.DataFrame.from_dict(text_counts,orient='index').reset_index()
        details.columns=['回复内容','人次']
        st.dataframe(details,width=800)

# 回复详情-第1，2，3次回复
def render_reply123(data:pd.DataFrame):
    if len(data)==0:
        return
    # 数据处理
    if(len(data)):
        reply=data.values.tolist()[0]
        pre=data.values.tolist()[1] if len(data)>1 else [0,0,0,0,0,0,0]
    else:
        reply=[0,0,0,0,0,0,0]
        pre=[0,0,0,0,0,0,0]
    # 绘图
    col1,col2,col3,col4=st.columns(4)
    col1.metric('总回复',f'{reply[0]}人次',f'{reply[0]-pre[0]}人次')

    col2.metric('第1次回复',f'{reply[1]}人次',f'{reply[1]-pre[1]}人次')
    col2.write(f'占比 {100*reply[4]:.2f}%(↑{100*(reply[4]-pre[4]):.2f}%)' if reply[4]>=pre[4] else f'占比 {100*reply[4]:.2f}%(↓{100*(pre[4]-reply[4]):.2f}%)')

    col3.metric('第2次回复',f'{reply[2]}人次',f'{reply[2]-pre[2]}人次')
    col3.write(f'占比 {100*reply[5]:.2f}%(↑{100*(reply[5]-pre[5]):.2f}%)' if reply[5]>=pre[5] else f'占比 {100*reply[5]:.2f}%(↓{100*(pre[5]-reply[5]):.2f}%)')
    
    col4.metric('第3次回复',f'{reply[3]}人次',f'{reply[3]-pre[3]}人次')
    col4.write(f'占比 {100*reply[6]:.2f}%(↑{100*(reply[6]-pre[6]):.2f}%)' if reply[6]>=pre[6] else f'占比 {100*reply[6]:.2f}%(↓{100*(pre[6]-reply[6]):.2f}%)')

# 回复数据特征表格
def render_reply_feature(data:pd.DataFrame):
    if len(data)==0:
        return
    d1=data.groupby(data['呼出开始时间'].dt.month).agg(
        转人工率=('按键回复',lambda x:x.str.contains('转人工').sum()/data.size), # 转人工率
        按键回复率=('按键回复',lambda x:x.str.contains('：').sum()/data.size), #   按键回复率
        人均重呼次数=('重呼次数',lambda x:x.values.sum()/data['个案编码'].nunique()), # 单人平均重呼次数
    )
    suc=data[data['呼出结果']=='呼出成功']
    d2=suc.groupby(suc['呼出开始时间'].dt.month).agg(
        成功接听人均重呼数=('重呼次数',lambda x:x.values.sum()/suc['个案编码'].nunique()),   # 单人成功接听所需重呼数
    )
    reply=pd.concat([d1[['转人工率','按键回复率','人均重呼次数']],d2],axis=1)
    reply.index.name='月份'

    # 定义样式
    style={
       '转人工率':'{0:.2%}',
       '按键回复率':'{0:.2%}',
       '人均重呼次数':'{0:.2f}次',
       '成功接听人均重呼数':'{0:.2f}次',
    }
    st.subheader('各月受种者回复数据特征')
    st.caption('单击列名查看升序或降序结果')
    st.dataframe(
        reply.style.format(style).background_gradient(
            subset=['转人工率','按键回复率'],
            cmap='Greens'
        )
        .highlight_min(
            subset=['人均重呼次数','成功接听人均重呼数'],
            props='background-color:pink'
        ),
        width=800,height=280,
        
    )

# 用户粘性分析 
def render_dau(data:pd.DataFrame):
    # 分别计算不同机构的日活
    d1=df[df['机构名称']=='明楼'].groupby(pd.Grouper(key='呼出开始时间',freq='ME')).agg(明楼=('呼出开始时间',lambda x: x.dt.date.nunique()))
    d2=df[df['机构名称']=='南码头社区卫生服务中心'].groupby(pd.Grouper(key='呼出开始时间',freq='ME')).agg(南码头社区卫生服务中心=('呼出开始时间',lambda x: x.dt.date.nunique()))
    d3=df[df['机构名称']=='北蔡社区卫生服务中心'].groupby(pd.Grouper(key='呼出开始时间',freq='ME')).agg(北蔡社区卫生服务中心=('呼出开始时间',lambda x: x.dt.date.nunique()))
    
    # 合并
    group=pd.concat([d1,d2,d3],axis=1)
    group=group.fillna(0)
    group.index=group.index.strftime('%Y-%m')
    group.index.name='月份'
    
    # 计算总数
    group['总计']=group['明楼'].values+group['南码头社区卫生服务中心'].values+group['北蔡社区卫生服务中心'].values
    print(group)
    style={
        '明楼':'{:.0f}天',
        '南码头社区卫生服务中心':'{:.0f}天',
        '北蔡社区卫生服务中心':'{:.0f}天',
        '总计':'{:.0f}天'
    }
    col1,col2=st.columns(2)
    with col1:
        st.dataframe(
            group.style.format(style).background_gradient(
                subset=['明楼','南码头社区卫生服务中心','北蔡社区卫生服务中心'],
                cmap='Reds'
            ).background_gradient(
                subset=['总计'],
                cmap='Greens'
            ),
            width=800
        )
    with col2:
        line=(
        Line(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND,width=f'{WIDTH*0.45}px',height=f'{HEIGHT*0.3}px'))
            .add_xaxis(group.index.astype('str').tolist())
            .add_yaxis('明楼',group['明楼'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
            .add_yaxis('南码头社区卫生服务中心',group['南码头社区卫生服务中心'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
            .add_yaxis('北蔡社区卫生服务中心',group['北蔡社区卫生服务中心'].values.tolist(),label_opts=opts.LabelOpts(is_show=False))
            .set_global_opts(
                tooltip_opts=opts.TooltipOpts(
                    trigger='axis',
                    axis_pointer_type='cross',
                ),
                datazoom_opts=opts.DataZoomOpts(),
                title_opts=opts.TitleOpts(title='各试点DAU统计'),
                legend_opts=opts.LegendOpts()
            )
        )
        # 组合绘图
        components.html(line.render_embed(),width=WIDTH*0.5,height=HEIGHT*0.35)

# -----------------------------------------------------------------------------------
# 页面配置
st.set_page_config(page_title="QRV呼出分析", layout="wide")
st.title("QRV呼出分析")

# 文件上传入口
files=st.file_uploader(accept_multiple_files=True,type={'xlsx'},label='上传各试点呼出数据文件(.xlsx)，可同时上传多个')
if len(files):
    df_input=[]
    for f in files:df_input.append(pd.read_excel(f))
    # 合并文件
    df=pd.concat(df_input,ignore_index=True)
    df['呼出开始时间']=pd.to_datetime(df['呼出开始时间'])

    # 按日期范围筛选数据入口
    data=data_filtering('range1')

    # 数据总览
    st.header("数据总览",help="QRV系统呼出情况基本数据")
    container_general=st.container()
    with container_general:
        # 频度选择
        freq1=render_freq_selectbox('general')
        # 2个子页面
        tab_total,tab_avg=st.tabs(["总数","平均"])
        # 数据查询
        total,avg=query_general(data,freq1)
        with tab_total:
            render_general_total(total)
        with tab_avg:
            render_general_avg(avg)

    # 呼出详情
    st.header("呼出详情")
    container_details=st.container()
    with container_details:
        # 频度选择
        freq2=render_freq_selectbox('detail')

        # 【变化曲线】板块
        with st.container():
            st.subheader("变化曲线")
            # 2个子页面
            tab_calls,tab_duration=st.tabs(["人次","时长"])
            with tab_calls:
                render_calls_detail(query_detail_fig(data,freq2))
            with tab_duration:
                col1,col2,col3=st.columns(3)
                with col1:
                    render_duration_detail(query_detail_fig(data,freq2))
                with col2:
                    render_task_detail(query_task_fig(data,freq2))
                with col3:
                    st.table(query_task_fig(data,freq2))

        # 【数据特征】板块
        with st.container():
            st.subheader("月度数据横向比较")
            col1,col2=st.columns([1.6,1])
            with col1:
                render_monthly(data)
            with col2:
                render_monthly_suc(data)

        # 成功和失败统计
        # 数据查询
        suc,fail=query_recall_reason(data,freq2)
        col_suc,col_fail=st.columns(2)
        with col_suc:
            with st.container():
                st.subheader("成功统计")
                st.caption('单击列名查看升序或降序结果')
                recall_rate=[float(rate.strip('%')) for rate in total['重呼率']]
                suc_detail={
                    '呼出成功(人次)':total['呼出成功'],
                    '成功率':total['成功率'],
                    '重呼任务(个)':(total['执行任务总数']*recall_rate).round(0),
                    '重呼率':total['重呼率']
                }
                st.dataframe(suc_detail,width=800)
                st.divider()
                render_suc_detail(suc)
            

        with col_fail:
            with st.container():
                st.subheader('失败统计')
                st.caption('单击列名查看升序或降序结果')
                fail_detail={
                    '呼出失败(人次)':total['呼出失败'],
                    '失败率':total['失败率']
                }
                st.dataframe(fail_detail,width=800)
                st.divider()
                render_fail_detail(fail)
                with st.expander('详细数据'):
                    st.dataframe(fail,width=700)

    # 受种者回复详情
    st.header('受种者详情')
    container_reply=st.container()
    with container_reply:
        freq3=render_freq_selectbox('reply')
        # 1，2，3轮回复人次及占比
        with st.container():
            render_reply123(query_reply_detail(data,freq3))
        # 回复内容分类
        st.divider()
        with st.container():
            col1,col2=st.columns(2)
            with col2:
                render_reply_classify(data)
            with col1:
                render_reply_feature(data)

    # 试点数据详情
    st.header('试点用户粘性统计')
    with st.container():
        dau_data=data_filtering('range2',filter_org=False)
        render_dau(dau_data)
