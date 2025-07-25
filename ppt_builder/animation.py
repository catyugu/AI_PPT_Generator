# 文件: test_animation_final_modular.py
# [终极模块化版 - 解决了重复创建和扩展性问题]

import logging
import os
from pptx import Presentation
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.enum.shapes import MSO_SHAPE

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_or_create_main_anim_container(slide_element):
    """
    [模块一：主厨] 获取或创建幻灯片的动画主容器。
    确保 <p:timing> 和主序列只被创建一次。
    :param slide_element: 幻灯片的 lxml 元素。
    :return: (可以附加动画效果的 <p:childTnLst> 节点, 主 <p:timing> 节点)
    """
    timing = slide_element.find(qn('p:timing'))
    if timing is None:
        logging.info("未找到 <p:timing> 节点，正在创建动画骨架...")
        # 使用经过验证的XML骨架，确保结构正确
        timing_xml_str = f"""
        <p:timing {nsdecls('p')}>
            <p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
            <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst/>
            </p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst/>
        </p:timing>"""
        timing = parse_xml(timing_xml_str)
        slide_element.append(timing)

    # 返回主序列中用于添加动画的那个 childTnLst 节点，以及timing节点本身
    main_anim_container = timing.find(qn('p:tnLst')).find(qn('p:par')).find(qn('p:cTn')).find(qn('p:childTnLst')).find(
        qn('p:seq')).find(qn('p:cTn')).find(qn('p:childTnLst'))
    return main_anim_container, timing


def build_fly_in_effect(shape_id, next_id, direction, duration_ms):
    """
    [模块二：菜谱] 构建一个独立的“飞入”动画效果XML片段。
    这个函数只关心效果本身，不关心容器。
    :return: 一个代表飞入效果的 <p:par> lxml 元素。
    """
    direction_map = {
        "fromBottom": {"preset": "4", "y_from": "1+#ppt_h/2", "x_from": "#ppt_x"},
        "fromLeft": {"preset": "8", "y_from": "#ppt_y", "x_from": "-#ppt_w"},
        "fromRight": {"preset": "2", "y_from": "#ppt_y", "x_from": "1"},
        "fromTop": {"preset": "1", "y_from": "-#ppt_h", "x_from": "#ppt_x"},
    }
    props = direction_map.get(direction, direction_map["fromBottom"])

    # 完整、精确地构建经过验证的XML结构
    effect_xml_str = """
    <p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cTn id="{id_3}" fill="hold">
            <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
            <p:childTnLst>
                <p:par>
                    <p:cTn id="{id_4}" fill="hold">
                        <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                        <p:childTnLst>
                            <p:par>
                                <p:cTn id="{id_5}" presetID="2" presetClass="entr" presetSubtype="{presetSubtype}" fill="hold" grpId="0" nodeType="clickEffect">
                                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                    <p:childTnLst>
                                        <p:set>
                                            <p:cBhvr>
                                                <p:cTn id="{id_6}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                                                <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                                                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                                            </p:cBhvr>
                                            <p:to><p:strVal val="visible"/></p:to>
                                        </p:set>
                                        <p:anim calcmode="lin" valueType="num">
                                            <p:cBhvr additive="base">
                                                <p:cTn id="{id_7}" dur="{duration}" fill="hold"/>
                                                <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                                                <p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>
                                            </p:cBhvr>
                                            <p:tavLst>
                                                <p:tav tm="0"><p:val><p:strVal val="{anim_x_from}"/></p:val></p:tav>
                                                <p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>
                                            </p:tavLst>
                                        </p:anim>
                                        <p:anim calcmode="lin" valueType="num">
                                            <p:cBhvr additive="base">
                                                <p:cTn id="{id_8}" dur="{duration}" fill="hold"/>
                                                <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                                                <p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>
                                            </p:cBhvr>
                                            <p:tavLst>
                                                <p:tav tm="0"><p:val><p:strVal val="{anim_y_from}"/></p:val></p:tav>
                                                <p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav>
                                            </p:tavLst>
                                        </p:anim>
                                    </p:childTnLst>
                                </p:cTn>
                            </p:par>
                        </p:childTnLst>
                    </p:cTn>
                </p:par>
            </p:childTnLst>
        </p:cTn>
    </p:par>
    """.format(
        id_3=str(next_id), id_4=str(next_id + 1), id_5=str(next_id + 2),
        id_6=str(next_id + 3), id_7=str(next_id + 4), id_8=str(next_id + 5),
        spid=shape_id, duration=str(duration_ms),
        presetSubtype=props["preset"], anim_x_from=props["x_from"], anim_y_from=props["y_from"]
    )
    return parse_xml(effect_xml_str)


def build_dissolve_in_effect(shape_id, next_id, duration_ms):
    """
    [新增的菜谱] 构建一个独立的“向内溶解”动画效果XML片段。
    :return: 一个代表向内溶解效果的 <p:par> lxml 元素。
    """
    # 这个XML模板精确复刻了您提供的slide1.xml中spid="5"的动画结构
    effect_xml_str = """
    <p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cTn id="{id_3}" fill="hold">
            <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
            <p:childTnLst>
                <p:par>
                    <p:cTn id="{id_4}" fill="hold">
                        <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                        <p:childTnLst>
                            <p:par>
                                <p:cTn id="{id_5}" presetID="9" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">
                                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                    <p:childTnLst>
                                        <p:set>
                                            <p:cBhvr>
                                                <p:cTn id="{id_6}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                                                <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                                                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                                            </p:cBhvr>
                                            <p:to><p:strVal val="visible"/></p:to>
                                        </p:set>
                                        <p:animEffect transition="in" filter="dissolve">
                                            <p:cBhvr>
                                                <p:cTn id="{id_7}" dur="{duration}"/>
                                                <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                                            </p:cBhvr>
                                        </p:animEffect>
                                    </p:childTnLst>
                                </p:cTn>
                            </p:par>
                        </p:childTnLst>
                    </p:cTn>
                </p:par>
            </p:childTnLst>
        </p:cTn>
    </p:par>
    """.format(
        id_3=str(next_id), id_4=str(next_id + 1), id_5=str(next_id + 2),
        id_6=str(next_id + 3), id_7=str(next_id + 4), spid=shape_id, duration=str(duration_ms)
    )
    return parse_xml(effect_xml_str)


def build_fade_in_effect(shape_id, next_id, duration_ms):
    """[Recipe 1: Fade In] Builds a fade-in effect, precisely replicating slide1.xml for spid="4"."""
    # This is a direct copy of the first animation block in your XML.
    effect_xml_str = """
    <p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cTn id="{id_3}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{id_4}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{id_5}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{id_6}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{id_7}" dur="{duration}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>""".format(id_3=str(next_id),id_4=str(next_id+1),id_5=str(next_id+2),id_6=str(next_id+3),id_7=str(next_id+4),spid=shape_id,duration=str(duration_ms))
    return parse_xml(effect_xml_str)

def build_fly_out_effect(shape_id, next_id, direction, duration_ms):
    """[修正后的菜谱四：飞出] 精确复刻 slide1.xml 中 spid="2" 的第二个动画。"""
    direction_map = {"toBottom": {"preset": "4", "y_to": "1+ppt_h/2", "x_to": "ppt_x"}, "toLeft": {"preset": "8", "y_to": "ppt_y", "x_to": "-ppt_w"}, "toRight": {"preset": "2", "y_to": "ppt_y", "x_to": "1"}, "toTop": {"preset": "1", "y_to": "-ppt_h", "x_to": "ppt_x"}}
    props = direction_map.get(direction, direction_map["toBottom"])
    effect_xml_str = """<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cTn id="{id_19}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{id_20}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{id_21}" presetID="2" presetClass="exit" presetSubtype="{presetSubtype}" fill="hold" grpId="1" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="{id_22}" dur="{duration}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="ppt_x"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="{anim_x_to}"/></p:val></p:tav></p:tavLst></p:anim><p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="{id_23}" dur="{duration}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="{anim_y_to}"/></p:val></p:tav></p:tavLst></p:anim><p:set><p:cBhvr><p:cTn id="{id_24}" dur="1" fill="hold"><p:stCondLst><p:cond delay="{hide_delay}"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>""".format(id_19=str(next_id),id_20=str(next_id+1),id_21=str(next_id+2),id_22=str(next_id+3),id_23=str(next_id+4),id_24=str(next_id+5),spid=shape_id,duration=str(duration_ms),presetSubtype=props["preset"],anim_x_to=props.get("x_to", "ppt_x"),anim_y_to=props.get("y_to", "ppt_y"),hide_delay=str(duration_ms-1))
    return parse_xml(effect_xml_str)


def build_fade_out_effect(shape_id, next_id, duration_ms):
    """[修正后的菜谱二：淡出] 精确复刻 slide1.xml 中 spid="4" 的第二个动画。"""
    effect_xml_str = """<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cTn id="{id_8}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{id_9}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{id_10}" presetID="10" presetClass="exit" presetSubtype="0" fill="hold" grpId="1" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="{id_11}" dur="{duration}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect><p:set><p:cBhvr><p:cTn id="{id_12}" dur="1" fill="hold"><p:stCondLst><p:cond delay="{hide_delay}"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>""".format(id_8=str(next_id),id_9=str(next_id+1),id_10=str(next_id+2),id_11=str(next_id+3),id_12=str(next_id+4),spid=shape_id,duration=str(duration_ms),hide_delay=str(duration_ms-1))
    return parse_xml(effect_xml_str)


# --- 模块三：动画系统的“服务员” ---
def add_animation(shape, effect: str, **kwargs):
    """为一个形状添加指定的动画效果。"""
    logging.info(f"为形状 ID:{shape.shape_id} 添加 '{effect}' 动画...")

    main_anim_container, timing = get_or_create_main_anim_container(shape.part._element)
    next_id = (len(list(main_anim_container)) * 10) + 3

    effect_fragment = None
    if effect == "flyIn":
        effect_fragment = build_fly_in_effect(str(shape.shape_id), next_id, kwargs.get("direction", "fromBottom"),
                                              kwargs.get("duration_ms", 500))
    elif effect == "fadeIn":  # 使用正确的名称
        effect_fragment = build_fade_in_effect(str(shape.shape_id), next_id, kwargs.get("duration_ms", 500))
    elif effect == "flyOut":
        effect_fragment = build_fly_out_effect(str(shape.shape_id), next_id, kwargs.get("direction", "toBottom"),
                                               kwargs.get("duration_ms", 500))
    elif effect == "fadeOut":
        effect_fragment = build_fade_out_effect(str(shape.shape_id), next_id, kwargs.get("duration_ms", 500))
    else:
        logging.warning(f"未知的动画效果类型: '{effect}'")
        return

    if effect_fragment is not None:
        main_anim_container.append(effect_fragment)
        bldLst = timing.find(qn('p:bldLst'))
        bldP = OxmlElement('p:bldP', {'spid': str(shape.shape_id), 'grpId': '0'})
        bldP.append(OxmlElement('p:bldSub'))
        bldLst.append(bldP)
        logging.info("动画效果已成功注入。")
