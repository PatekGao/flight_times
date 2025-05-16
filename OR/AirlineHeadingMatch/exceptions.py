"""
管理航司和航向的例外配置
"""

# 到达航班例外配置
ARR_INT_EXCEPTIONS = {
    '国航集团': ['LONGKOU'],
    '南航集团': ['XISHUI'],
    '海航集团': ['LONGKOU']
}

ARR_DOM_EXCEPTIONS = {
    # 国内到达航线的例外配置
}

# 出发航班例外配置
DEP_INT_EXCEPTIONS = {
    '国航集团': ['XISHUI'],
    '南航集团': ['XISHUI'],
    '海航集团': ['XISHUI']
}

DEP_DOM_EXCEPTIONS = {
    # 国内出发航线的例外配置
}

def get_valid_airline_heading_pairs(market_type, direction):
    """
    根据市场类型和方向获取有效的航司-航向对
    
    Args:
        market_type (str): 市场类型，如 'INT' 表示国际航线，'DOM' 表示国内航线
        direction (str): 方向，'ARR' 表示到达，'DEP' 表示出发
        
    Returns:
        set: 包含 (airline, heading) 元组的集合
    """
    valid_pairs = set()
    
    if direction == 'ARR':
        if market_type == 'INT':
            exceptions = ARR_INT_EXCEPTIONS
        else:
            exceptions = ARR_DOM_EXCEPTIONS
    else:  # direction == 'DEP'
        if market_type == 'INT':
            exceptions = DEP_INT_EXCEPTIONS
        else:
            exceptions = DEP_DOM_EXCEPTIONS
            
    for airline, headings in exceptions.items():
        for heading in headings:
            valid_pairs.add((airline, heading))
                
    return valid_pairs 