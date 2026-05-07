import os
import re
import math
import glob
import copy
import subprocess
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image as PILImage
import fitz

try:
    import yaml
except ImportError as e:
    raise ImportError("缺少 PyYAML 依赖，请先安装：pip install pyyaml") from e


# 高级版完整 schema 模板：当用户未提供 report_schema.yaml 时，quickstart 模式将以此为基础展开。
FULL_SCHEMA_TEMPLATE = {'schema_version': '2.0',
 'project': {'name': 'MintPy Word Auto Report', 'language': 'zh-CN', 'description': '基于 MintPy 结果自动生成 Word 报告'},
 'runtime_paths': {'logo_path': '/home/liyuxuan/penguin/pic2word/pic/logo.png',
                   'base_dir': '/home/liyuxuan/penguin/data/Galapagos_s1/mintpy',
                   'pic_dir': '{runtime_paths.base_dir}/pic',
                   'output_dir': '/home/liyuxuan/penguin/pic2word/vision_2/word',
                   'velocity_file': '{runtime_paths.base_dir}/velocity.h5',
                   'cfg_file': '{runtime_paths.pic_dir}/smallbaselineApp.cfg',
                   'rms_txt': '{runtime_paths.pic_dir}/rms_timeseriesResidual_ramp.txt',
                   'reference_date_txt': '{runtime_paths.pic_dir}/reference_date.txt',
                   'exclude_date_txt': '{runtime_paths.pic_dir}/exclude_date.txt'},
 'defaults': {'font_cn': '仿宋',
              'font_en': 'Times New Roman',
              'figure_width_cm': 14.0,
              'figure_max_height_cm': 14.0,
              'table_style': 'Table Grid',
              'missing_asset_policy': 'warn',
              'auto_page_break_after_chapter': True,
              'pdf_render_dpi': 300,
              'figure_numbering_scope': 'chapter',
              'table_numbering_scope': 'chapter'},
 'data_sources': {'mintpy_meta': {'type': 'mintpy_attribute', 'path': '{runtime_paths.velocity_file}'},
                  'obs_dates': {'type': 'text_file', 'path': '{runtime_paths.rms_txt}', 'parser': 'date_list'},
                  'mintpy_cfg': {'type': 'cfg_file', 'path': '{runtime_paths.cfg_file}'},
                  'reference_dates': {'type': 'text_file',
                                      'path': '{runtime_paths.reference_date_txt}',
                                      'parser': 'inline_text'},
                  'exclude_dates': {'type': 'text_file',
                                    'path': '{runtime_paths.exclude_date_txt}',
                                    'parser': 'inline_text'},
                  'ts_method': {'type': 'interactive', 'prompt': '>>> 请输入时序估计方法 (默认: SBAS): ', 'default': 'SBAS'},
                  'analysis_text': {'type': 'interactive', 'prompt': '>>> 请在此处输入分析内容: ', 'default': '（请补充解译分析内容）'},
                  'poi_figures': {'type': 'glob',
                                  'pattern': '{runtime_paths.pic_dir}/final_figure*.png',
                                  'sort': 'natural'}},
 'assets': {'network_figure': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/network.pdf'},
            'coherence_history': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/coherenceHistory.pdf'},
            'coherence_matrix': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/coherenceMatrix.pdf'},
            'avg_spatial_coh': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/geo_avgSpatialCoh.png'},
            'temporal_coh': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/geo_temporalCoherence.png'},
            'mask_conn_comp': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/maskConnComp.png'},
            'avg_phase_velocity': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/avgPhaseVelocity.png'},
            'unwrap_error': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/numTriNonzeroIntAmbiguity.png'},
            'geometry_radar': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/geometryRadar.png'},
            'residual_rms': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/rms_timeseriesResidual_ramp.pdf'},
            'velocity_map': {'type': 'figure',
                             'candidates': ['{runtime_paths.pic_dir}/velocity_DEM.png',
                                            '{runtime_paths.pic_dir}/velocity.png']},
            'velocity_std': {'type': 'figure', 'path': '{runtime_paths.pic_dir}/velocity_std_plot.png'},
            'poi_series': {'type': 'figure_group', 'source': 'poi_figures'}},
 'templates': {'headings': {'cover_title': '基于{user_inputs.data_source}数据的{user_inputs.region_name}地区\n'
                                           '时序InSAR地表形变处理与分析报告'},
               'captions': {'network': '图{figure_no} '
                                       '干涉图网络。图中节点表示SAR影像的获取日期，连线表示用于生成干涉图的影像对。横轴为时间，纵轴为相对于参考影像的垂直基线（单位：米）。连线的颜色映射表示各干涉对的平均空间相干性数值，取值范围为0.2至1.0。',
                            'coherence_history': '图{figure_no} '
                                                 '相干性极值统计。图中横轴表示SAR影像的获取日期，纵轴表示平均空间相干性系数，取值范围为0至1。对于每一个观测日期，柱状图展示了与该影像相关联的所有干涉对的相干性极值统计。',
                            'coherence_matrix': '图{figure_no} 相干性矩阵。图中横轴与纵轴均表示SAR影像按获取时间先后排序的索引编号。',
                            'avg_spatial_coh': '图{figure_no} 地理编码后的平均空间相干性图。',
                            'temporal_coh': '图{figure_no} 地理编码后时间相干性图。',
                            'mask_conn_comp': '图{figure_no} 相位解缠连通分量观测掩膜。',
                            'avg_phase_velocity': '图{figure_no} 平均相位堆栈速率图。',
                            'unwrap_error': '图{figure_no} 相位解缠误差评估图。',
                            'geometry_radar': '图{figure_no} 雷达几何参数与辅助数据图。',
                            'residual_rms': '图{figure_no} 残余相位 RMS 可视化结果。',
                            'velocity_map': '图{figure_no} 形变速率图',
                            'velocity_std': '图{figure_no} VelocityStd',
                            'poi_timeseries': '图{figure_no} 感兴趣目标点{item_index}的时序形变图'},
               'text': {'reference_dates': '参考日期：{value}', 'exclude_dates': '剔除日期：{value}'}},
 'report_structure': [{'type': 'cover',
                       'params': {'data_source': 'interactive',
                                  'region_name': 'interactive',
                                  'lead_author': 'interactive',
                                  'logo_path': '{runtime_paths.logo_path}',
                                  'output_name_template': '{user_inputs.region_name}InSAR处理报告.docx'},
                       'default_data_source': 'Sentinel-1',
                       'default_region_name': '北京',
                       'default_lead_author': 'yourname'},
                      {'type': 'toc'},
                      {'type': 'chapter',
                       'id': 'chapter_1',
                       'title': '1. 数据情况',
                       'children': [{'type': 'metadata_table',
                                     'source': 'mintpy_meta',
                                     'fields': [['卫 星', 'satellite'],
                                                ['波 段', 'band'],
                                                ['波 长', 'wavelength'],
                                                ['成像模式', 'mode'],
                                                ['swath_Num', 'swath_Num'],
                                                ['空间分辨率', 'resolution'],
                                                ['相对轨道号', 'track'],
                                                ['升降轨', 'direction'],
                                                ['中心入射角', 'inc_angle'],
                                                ['中心方位角', 'az_angle'],
                                                ['极化方式', 'pol'],
                                                ['开始时间', 'start_time'],
                                                ['结束时间', 'end_time'],
                                                ['干涉对数量', 'num_ifgram'],
                                                ['经纬度范围', 'bbox']]},
                                    {'type': 'table',
                                     'title': '表{table_no} 选用SAR数据日期',
                                     'source': 'obs_dates',
                                     'renderer': 'date_grid',
                                     'columns': 4}]},
                      {'type': 'chapter',
                       'id': 'chapter_2',
                       'title': '2. 处理参数设置',
                       'children': [{'type': 'software_versions',
                                     'label': '处理软件',
                                     'platform_source': 'mintpy_meta',
                                     'stack_processor': 'auto'},
                                    {'type': 'kv_line', 'label': '时序估计方法', 'source': 'ts_method'},
                                    {'type': 'cfg_table',
                                     'title': 'MintPy 参数配置',
                                     'source': 'mintpy_cfg',
                                     'filter': {'exclude_values': ['auto', 'no']}}]},
                      {'type': 'chapter',
                       'id': 'chapter_3',
                       'title': '3. 干涉数据质量评估',
                       'children': [{'type': 'section',
                                     'id': 'chapter_3_1',
                                     'title': '3.1 干涉图网络',
                                     'children': [{'type': 'figure',
                                                   'asset': 'network_figure',
                                                   'caption': '{templates.captions.network}'},
                                                  {'type': 'figure',
                                                   'asset': 'coherence_history',
                                                   'caption': '{templates.captions.coherence_history}'},
                                                  {'type': 'figure',
                                                   'asset': 'coherence_matrix',
                                                   'caption': '{templates.captions.coherence_matrix}'}]},
                                    {'type': 'section',
                                     'id': 'chapter_3_2',
                                     'title': '3.2 相干性',
                                     'children': [{'type': 'figure',
                                                   'asset': 'avg_spatial_coh',
                                                   'caption': '{templates.captions.avg_spatial_coh}'},
                                                  {'type': 'figure',
                                                   'asset': 'temporal_coh',
                                                   'caption': '{templates.captions.temporal_coh}'},
                                                  {'type': 'figure',
                                                   'asset': 'mask_conn_comp',
                                                   'caption': '{templates.captions.mask_conn_comp}'}]},
                                    {'type': 'section',
                                     'id': 'chapter_3_3',
                                     'title': '3.3 相位堆栈',
                                     'children': [{'type': 'figure',
                                                   'asset': 'avg_phase_velocity',
                                                   'caption': '{templates.captions.avg_phase_velocity}',
                                                   'max_height_cm': 12}]},
                                    {'type': 'section',
                                     'id': 'chapter_3_4',
                                     'title': '3.4 相位解缠误差',
                                     'children': [{'type': 'figure',
                                                   'asset': 'unwrap_error',
                                                   'caption': '{templates.captions.unwrap_error}',
                                                   'max_height_cm': 12}]},
                                    {'type': 'section',
                                     'id': 'chapter_3_5',
                                     'title': '3.5 几何数据',
                                     'children': [{'type': 'figure',
                                                   'asset': 'geometry_radar',
                                                   'caption': '{templates.captions.geometry_radar}',
                                                   'max_height_cm': 12}]},
                                    {'type': 'section',
                                     'id': 'chapter_3_6',
                                     'title': '3.6 残余相位 RMS',
                                     'children': [{'type': 'figure',
                                                   'asset': 'residual_rms',
                                                   'caption': '{templates.captions.residual_rms}',
                                                   'max_height_cm': 12},
                                                  {'type': 'text_file_line',
                                                   'source': 'reference_dates',
                                                   'template': '{templates.text.reference_dates}'},
                                                  {'type': 'text_file_line',
                                                   'source': 'exclude_dates',
                                                   'template': '{templates.text.exclude_dates}',
                                                   'space_after_pt': 12}]}]},
                      {'type': 'chapter',
                       'id': 'chapter_4',
                       'title': '4. 形变结果',
                       'children': [{'type': 'section',
                                     'id': 'chapter_4_1',
                                     'title': '4.1 形变速率图',
                                     'children': [{'type': 'conditional',
                                                   'when': {'asset_exists': 'velocity_map'},
                                                   'then': [{'type': 'figure',
                                                             'asset': 'velocity_map',
                                                             'caption': '{templates.captions.velocity_map}'}],
                                                   'else': [{'type': 'paragraph',
                                                             'text': '[图片缺失: velocity_DEM.png 或 velocity.png]'}]}]},
                                    {'type': 'section',
                                     'id': 'chapter_4_2',
                                     'title': '4.2 VelocityStd',
                                     'children': [{'type': 'conditional',
                                                   'when': {'asset_exists': 'velocity_std'},
                                                   'then': [{'type': 'figure',
                                                             'asset': 'velocity_std',
                                                             'caption': '{templates.captions.velocity_std}'}],
                                                   'else': [{'type': 'paragraph',
                                                             'text': '[图片缺失: velocity_std_plot.png]'}]}]},
                                    {'type': 'section',
                                     'id': 'chapter_4_3',
                                     'title': '4.3 感兴趣目标点的时序形变图',
                                     'children': [{'type': 'conditional',
                                                   'when': {'source_nonempty': 'poi_figures'},
                                                   'then': [{'type': 'repeat',
                                                             'source': 'poi_figures',
                                                             'as': 'item',
                                                             'children': [{'type': 'figure',
                                                                           'path': '{item.path}',
                                                                           'caption': '{templates.captions.poi_timeseries}'}]}],
                                                   'else': [{'type': 'paragraph', 'text': '（未检测到感兴趣目标点图件）'}]}]}]},
                      {'type': 'chapter',
                       'id': 'chapter_5',
                       'title': '5. 解译分析',
                       'children': [{'type': 'interactive_text',
                                     'source': 'analysis_text',
                                     'line_spacing': 1.5,
                                     'first_line_indent_pt': 24,
                                     'font_size_pt': 12,
                                     'font_name_cn': '仿宋'}]}]}

# 兼容旧代码中的命名
DEFAULT_SCHEMA = copy.deepcopy(FULL_SCHEMA_TEMPLATE)

_SCHEMA_CACHE = {}
_PLACEHOLDER_RE = re.compile(r"\{([^{}]+)\}")


def deep_merge(base, override):
    result = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def get_default_config_candidates():
    if "__file__" in globals():
        module_dir = os.path.dirname(os.path.abspath(__file__))
    else:
        module_dir = os.getcwd()

    # 方案四：优先 quickstart，找不到再用高级 schema
    return [
        os.path.join(module_dir, "report_quickstart.yaml"),
        os.path.join(module_dir, "report_schema.yaml"),
        os.path.join(module_dir, "captions.yaml"),
        os.path.join(os.getcwd(), "report_quickstart.yaml"),
        os.path.join(os.getcwd(), "report_schema.yaml"),
        os.path.join(os.getcwd(), "captions.yaml"),
    ]


def validate_schema(cfg):
    required_keys = [
        "schema_version",
        "runtime_paths",
        "defaults",
        "data_sources",
        "assets",
        "templates",
        "report_structure",
    ]
    for key in required_keys:
        if key not in cfg:
            raise ValueError(f"YAML 缺少顶层字段: {key}")

    if not isinstance(cfg["report_structure"], list):
        raise ValueError("report_structure 必须为列表。")


def load_yaml_dict(path):
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if not isinstance(data, dict):
        raise ValueError("YAML 顶层结构必须为字典。")
    return data


def is_quickstart_config(cfg):
    return (
        isinstance(cfg, dict)
        and (
            str(cfg.get("schema_version", "")).startswith("quickstart")
            or "paths" in cfg
            or "cover_defaults" in cfg
            or "report_options" in cfg
        )
    )


def is_full_schema_config(cfg):
    return (
        isinstance(cfg, dict)
        and "report_structure" in cfg
        and "runtime_paths" in cfg
        and "data_sources" in cfg
        and "assets" in cfg
    )


def set_interactive_default(full_cfg, source_name, default_value):
    if default_value is None:
        return
    source_cfg = full_cfg.get("data_sources", {}).get(source_name)
    if isinstance(source_cfg, dict) and source_cfg.get("type") == "interactive":
        source_cfg["default"] = default_value


def find_block_by_id(blocks, block_id):
    if not isinstance(blocks, list):
        return None
    for block in blocks:
        if isinstance(block, dict):
            if block.get("id") == block_id:
                return block
            found = find_block_by_id(block.get("children", []), block_id)
            if found is not None:
                return found
    return None


def apply_titles_override(full_cfg, titles_override):
    if not isinstance(titles_override, dict):
        return
    for block_id, new_title in titles_override.items():
        block = find_block_by_id(full_cfg.get("report_structure", []), block_id)
        if block is not None and isinstance(new_title, str) and new_title.strip():
            block["title"] = new_title.strip()


def remove_chapter(full_cfg, chapter_id):
    report_structure = full_cfg.get("report_structure", [])
    full_cfg["report_structure"] = [
        block for block in report_structure
        if not (isinstance(block, dict) and block.get("type") == "chapter" and block.get("id") == chapter_id)
    ]


def remove_section(full_cfg, chapter_id, section_id):
    chapter = find_block_by_id(full_cfg.get("report_structure", []), chapter_id)
    if not chapter:
        return
    children = chapter.get("children", [])
    chapter["children"] = [
        block for block in children
        if not (isinstance(block, dict) and block.get("id") == section_id)
    ]


def load_base_full_schema_from_quickstart(quickstart_path=None):
    # quickstart 与高级 schema 共存：如果 quickstart 同目录下有 report_schema.yaml，则优先用它作高级模板；
    # 否则退回 utils.py 内置模板。
    candidates = []
    if quickstart_path:
        quick_dir = os.path.dirname(os.path.abspath(quickstart_path))
        candidates.append(os.path.join(quick_dir, "report_schema.yaml"))
    if "__file__" in globals():
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "report_schema.yaml"))
    candidates.append(os.path.join(os.getcwd(), "report_schema.yaml"))

    for path in candidates:
        if path and os.path.exists(path):
            try:
                data = load_yaml_dict(path)
                if is_full_schema_config(data):
                    return deep_merge(FULL_SCHEMA_TEMPLATE, data), path
            except Exception as e:
                print(f"[警告] 读取高级 schema 失败，将回退到内置模板: {e}")

    return copy.deepcopy(FULL_SCHEMA_TEMPLATE), None


def build_full_schema_from_quickstart(quick_cfg, quickstart_path=None):
    full_cfg, schema_source = load_base_full_schema_from_quickstart(quickstart_path)

    if schema_source:
        print(f"[提示] quickstart 模式使用高级模板: {schema_source}")
    else:
        print("[提示] quickstart 模式未检测到外部 report_schema.yaml，使用 utils.py 内置完整模板。")

    # project
    project_name = quick_cfg.get("project_name")
    if isinstance(project_name, str) and project_name.strip():
        full_cfg.setdefault("project", {})["name"] = project_name.strip()

    # 路径
    paths = quick_cfg.get("paths", {})
    runtime_paths = full_cfg.setdefault("runtime_paths", {})
    if isinstance(paths, dict):
        if paths.get("logo_path"):
            runtime_paths["logo_path"] = paths["logo_path"]
        if paths.get("base_dir"):
            runtime_paths["base_dir"] = paths["base_dir"]
        if paths.get("output_dir"):
            runtime_paths["output_dir"] = paths["output_dir"]

    # 重新绑定依赖 base_dir / pic_dir 的默认派生路径
    runtime_paths["pic_dir"] = "{runtime_paths.base_dir}/pic"
    runtime_paths["velocity_file"] = "{runtime_paths.base_dir}/velocity.h5"
    runtime_paths["cfg_file"] = "{runtime_paths.pic_dir}/smallbaselineApp.cfg"
    runtime_paths["rms_txt"] = "{runtime_paths.pic_dir}/rms_timeseriesResidual_ramp.txt"
    runtime_paths["reference_date_txt"] = "{runtime_paths.pic_dir}/reference_date.txt"
    runtime_paths["exclude_date_txt"] = "{runtime_paths.pic_dir}/exclude_date.txt"

    # 封面默认值
    cover_defaults = quick_cfg.get("cover_defaults", {})
    cover_block = None
    for block in full_cfg.get("report_structure", []):
        if isinstance(block, dict) and block.get("type") == "cover":
            cover_block = block
            break

    if cover_block:
        if isinstance(cover_defaults, dict):
            if "data_source" in cover_defaults:
                cover_block["default_data_source"] = cover_defaults["data_source"]
            if "region_name" in cover_defaults:
                cover_block["default_region_name"] = cover_defaults["region_name"]
            if "lead_author" in cover_defaults:
                cover_block["default_lead_author"] = cover_defaults["lead_author"]
            if "output_name_template" in cover_defaults:
                cover_block.setdefault("params", {})["output_name_template"] = cover_defaults["output_name_template"]

    # 处理参数与解译分析默认值
    processing_defaults = quick_cfg.get("processing_defaults", {})
    if isinstance(processing_defaults, dict):
        set_interactive_default(full_cfg, "ts_method", processing_defaults.get("ts_method"))
        set_interactive_default(full_cfg, "analysis_text", processing_defaults.get("analysis_text"))

    # 报告选项
    report_options = quick_cfg.get("report_options", {})
    if isinstance(report_options, dict):
        if "auto_page_break_after_chapter" in report_options:
            full_cfg.setdefault("defaults", {})["auto_page_break_after_chapter"] = bool(report_options["auto_page_break_after_chapter"])

        if report_options.get("include_chapter_3") is False:
            remove_chapter(full_cfg, "chapter_3")
        if report_options.get("include_chapter_4") is False:
            remove_chapter(full_cfg, "chapter_4")
        if report_options.get("include_chapter_5") is False:
            remove_chapter(full_cfg, "chapter_5")

        include_poi_section = str(report_options.get("include_poi_section", "auto")).lower()
        if include_poi_section in ("false", "0", "no", "n"):
            remove_section(full_cfg, "chapter_4", "chapter_4_3")

    # 图注覆盖
    captions_override = quick_cfg.get("captions_override", {})
    templates = full_cfg.setdefault("templates", {}).setdefault("captions", {})
    if isinstance(captions_override, dict):
        for key, value in captions_override.items():
            if isinstance(value, str) and value.strip():
                templates[key] = value.strip()

    # 标题覆盖
    apply_titles_override(full_cfg, quick_cfg.get("titles_override", {}))

    # 允许通过 quickstart 传递少量 defaults 覆盖
    defaults_override = quick_cfg.get("defaults_override", {})
    if isinstance(defaults_override, dict):
        full_cfg["defaults"] = deep_merge(full_cfg.get("defaults", {}), defaults_override)

    full_cfg["_config_mode"] = "quickstart"
    full_cfg["_quickstart_path"] = quickstart_path
    full_cfg["_base_schema_source"] = schema_source
    return full_cfg


def load_report_config(yaml_path=None, force_reload=False):
    candidate_paths = [yaml_path] if yaml_path else get_default_config_candidates()

    resolved_yaml_path = None
    for path in candidate_paths:
        if path and os.path.exists(path):
            resolved_yaml_path = path
            break

    cache_key = resolved_yaml_path or "__default__"
    if not force_reload and cache_key in _SCHEMA_CACHE:
        return _SCHEMA_CACHE[cache_key]

    if resolved_yaml_path:
        raw_cfg = load_yaml_dict(resolved_yaml_path)
        if is_quickstart_config(raw_cfg):
            print(f"[提示] 已加载 quickstart 配置文件: {resolved_yaml_path}")
            cfg = build_full_schema_from_quickstart(raw_cfg, quickstart_path=resolved_yaml_path)
        elif is_full_schema_config(raw_cfg):
            print(f"[提示] 已加载高级 schema 配置文件: {resolved_yaml_path}")
            cfg = deep_merge(FULL_SCHEMA_TEMPLATE, raw_cfg)
            cfg["_config_mode"] = "advanced"
        else:
            raise ValueError("无法识别该 YAML 类型：既不是 quickstart，也不是完整 report_schema。")
    else:
        print("[提示] 未检测到配置文件，将使用 utils.py 内置完整模板。")
        cfg = copy.deepcopy(FULL_SCHEMA_TEMPLATE)
        cfg["_config_mode"] = "builtin"

    validate_schema(cfg)
    cfg["_yaml_path"] = resolved_yaml_path
    _SCHEMA_CACHE[cache_key] = cfg
    return cfg



def get_by_dotted_path(data, path, default=None):
    if path is None:
        return default
    current = data
    for part in str(path).split("."):
        if isinstance(current, dict) and part in current:
            current = current[part]
        else:
            return default
    return current


def resolve_string_template(text, context):
    if not isinstance(text, str):
        return text

    def replacer(match):
        expr = match.group(1).strip()
        value = get_by_dotted_path(context, expr, default=match.group(0))
        if value is None:
            return ""
        return str(value)

    previous = None
    current = text
    max_rounds = 10
    rounds = 0
    while previous != current and rounds < max_rounds:
        previous = current
        current = _PLACEHOLDER_RE.sub(replacer, current)
        rounds += 1
    return current


def resolve_object(obj, context):
    if isinstance(obj, str):
        return resolve_string_template(obj, context)
    if isinstance(obj, list):
        return [resolve_object(x, context) for x in obj]
    if isinstance(obj, dict):
        return {k: resolve_object(v, context) for k, v in obj.items()}
    return obj


def set_font(run, font_size_pt, is_bold=False, font_name_cn="仿宋"):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name_cn)
    run.font.size = Pt(font_size_pt)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)


def set_cell_bottom_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), "000000")
    tcBorders.append(bottom)


def create_toc_element(doc):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()

    fldChar_begin = OxmlElement("w:fldChar")
    fldChar_begin.set(qn("w:fldCharType"), "begin")
    run._element.append(fldChar_begin)

    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = r'TOC \o "1-3" \h \z \u'
    run._element.append(instrText)

    fldChar_separate = OxmlElement("w:fldChar")
    fldChar_separate.set(qn("w:fldCharType"), "separate")
    run._element.append(fldChar_separate)

    run_display = paragraph.add_run()
    run_display.text = "（目录生成区：生成文档后，请右键点击此处 -> 更新域 -> 更新整个目录）"
    set_font(run_display, 10, is_bold=False)

    fldChar_end = OxmlElement("w:fldChar")
    fldChar_end.set(qn("w:fldCharType"), "end")
    run_display._element.append(fldChar_end)


def convert_pdf_to_png(pdf_path, dpi=300):
    if not os.path.exists(pdf_path):
        print(f"[警告] 文件不存在: {pdf_path}")
        return None

    png_path = pdf_path.replace(".pdf", ".png")
    if os.path.exists(png_path) and os.path.getmtime(png_path) > os.path.getmtime(pdf_path):
        return png_path

    try:
        doc_pdf = fitz.open(pdf_path)
        page = doc_pdf.load_page(0)
        pix = page.get_pixmap(dpi=dpi)
        pix.save(png_path)
        doc_pdf.close()
        return png_path
    except Exception as e:
        print(f"[错误] PDF 转图片失败 ({pdf_path}): {e}")
        return None


def read_text_file_safely(*candidate_paths):
    encodings = ["utf-8", "utf-8-sig", "gbk", "gb18030"]
    for path in candidate_paths:
        if not path or not os.path.exists(path):
            continue
        for enc in encodings:
            try:
                with open(path, "r", encoding=enc) as f:
                    lines = [line.strip() for line in f.readlines() if line.strip()]
                return "; ".join(lines)
            except UnicodeDecodeError:
                continue
            except Exception as e:
                print(f"[警告] 读取文件失败 ({path}): {e}")
                return None
    return None


def read_date_list_from_txt(file_path):
    date_list = []
    demo_text = """
    20230109    0.019
    20230121    0.022
    20230202    0.018
    20230214    0.017
    20230226    0.015
    """
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
        except Exception as e:
            print(f"[警告] 读取 TXT 文件失败: {e}")
            return []
    else:
        print(f"[提示] 文件 {file_path} 不存在，使用模拟数据演示表格生成。")
        lines = demo_text.strip().split("\n")

    for line in lines:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split()
        if len(parts) >= 1:
            date_list.append(parts[0])
    return date_list


def parse_sar_metadata(atr):
    meta = {}

    plat = atr.get("PLATFORM", "").lower()
    if "sen" in plat:
        meta["satellite"] = "Sentinel-1"
    elif "alos" in plat:
        meta["satellite"] = "ALOS"
    else:
        meta["satellite"] = atr.get("PLATFORM", "Unknown")

    wl = float(atr.get("WAVELENGTH", 0.056))
    meta["wavelength"] = f"{wl * 100:.1f} cm"
    if 0.024 <= wl <= 0.038:
        meta["band"] = "X"
    elif 0.038 <= wl <= 0.075:
        meta["band"] = "C"
    elif 0.150 <= wl <= 0.300:
        meta["band"] = "L"
    else:
        meta["band"] = "Unknown"

    meta["mode"] = atr.get("beam_mode", atr.get("ACQUISITION_MODE", "IW"))
    meta["swath_Num"] = atr.get("beam_swath", atr.get("beam_swath", "N/A"))

    rg_size = float(atr.get("rangePixelSize", atr.get("RANGE_PIXEL_SIZE", 0)))
    az_size = float(atr.get("azimuthPixelSize", atr.get("AZIMUTH_PIXEL_SIZE", 0)))
    unit = "米" if rg_size > 1 else "度"
    meta["resolution"] = f"{rg_size:.1f} {unit} (距离向) × {az_size:.1f} {unit} (方位向)"
    meta["track"] = atr.get("trackNumber", "N/A")

    orbit_dir = atr.get("ORBIT_DIRECTION", atr.get("passDirection", "")).upper()
    if "DESC" in orbit_dir:
        meta["direction"] = "降轨 (Descending)"
    elif "ASC" in orbit_dir:
        meta["direction"] = "升轨 (Ascending)"
    else:
        heading = float(atr.get("HEADING", 0))
        meta["direction"] = "降轨 (Descending)" if -180 <= heading <= -90 else "升轨 (Ascending)"

    meta["inc_angle"] = f"{float(atr.get('CENTER_INCIDENCE_ANGLE', 0)):.2f}°"
    meta["az_angle"] = f"{float(atr.get('HEADING', 0)):.2f}°"
    meta["pol"] = atr.get("POLARIZATION", atr.get("polarization", "VV"))

    s_date = atr.get("START_DATE", "")
    e_date = atr.get("END_DATE", "")
    if len(s_date) == 8:
        s_date = f"{s_date[:4]}-{s_date[4:6]}-{s_date[6:8]}"
    if len(e_date) == 8:
        e_date = f"{e_date[:4]}-{e_date[4:6]}-{e_date[6:8]}"
    meta["start_time"] = s_date
    meta["end_time"] = e_date

    try:
        lats = [float(atr.get(f"LAT_REF{i}")) for i in range(1, 5)]
        lons = [float(atr.get(f"LON_REF{i}")) for i in range(1, 5)]
        meta["bbox"] = f"{min(lats):.2f}°N ~ {max(lats):.2f}°N, {min(lons):.2f}°E ~ {max(lons):.2f}°E"
    except Exception:
        meta["bbox"] = "N/A"

    meta["num_ifgram"] = atr.get("mintpy.networkInversion.numIfgram", "N/A")
    return meta


def parse_cfg_file(cfg_path):
    params = []
    if not os.path.exists(cfg_path):
        return params

    try:
        with open(cfg_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        for line in lines:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, raw_val = line.split("=", 1)
                key = key.strip()
                val_clean = raw_val.split("#")[0].strip().strip("'").strip('"')
                if val_clean:
                    params.append((key, val_clean))
    except Exception as e:
        print(f"[错误] 解析配置文件 {os.path.basename(cfg_path)} 失败: {e}")
    return params


def get_software_versions():
    versions = {"ISCE": None, "MintPy": "Unknown"}

    try:
        import mintpy
        if hasattr(mintpy, "__version__"):
            versions["MintPy"] = mintpy.__version__
        elif hasattr(mintpy, "version"):
            versions["MintPy"] = mintpy.version.release_version
    except Exception:
        versions["MintPy"] = "未安装"

    try:
        result = subprocess.run("isce_version.py", capture_output=True, text=True, shell=True)
        if result.returncode == 0 and result.stdout.strip():
            out = result.stdout.strip()
            versions["ISCE"] = out.split()[1] if len(out.split()) > 1 else out
    except Exception:
        pass

    if not versions["ISCE"]:
        versions["ISCE"] = "Unknown"
    return versions


def add_text_line(doc, text, font_size_pt=10.5, font_name_cn="仿宋", space_after_pt=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(space_after_pt)
    run = p.add_run(text)
    set_font(run, font_size_pt, is_bold=False, font_name_cn=font_name_cn)


def add_figure_with_caption(doc, image_path, caption_text, target_width_cm=14.0, max_height_cm=14.0, pdf_render_dpi=300):
    final_img_path = image_path
    if image_path.lower().endswith(".pdf"):
        converted_path = convert_pdf_to_png(image_path, dpi=pdf_render_dpi)
        if converted_path:
            final_img_path = converted_path
        else:
            p_err = doc.add_paragraph(f"[图片缺失: {os.path.basename(image_path)}]")
            p_err.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p_err.runs:
                set_font(p_err.runs[0], 10.5, False, "Times New Roman")
            return

    if not os.path.exists(final_img_path):
        p_err = doc.add_paragraph(f"[图片缺失: {os.path.basename(final_img_path)}]")
        p_err.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p_err.runs:
            set_font(p_err.runs[0], 10.5, False, "Times New Roman")
        return

    try:
        with PILImage.open(final_img_path) as img:
            w_px, h_px = img.size
            aspect_ratio = w_px / h_px if h_px else 1.0
            height_if_width_fixed = target_width_cm / aspect_ratio

            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_img = p_img.add_run()

            if height_if_width_fixed > max_height_cm:
                run_img.add_picture(final_img_path, height=Cm(max_height_cm))
            else:
                run_img.add_picture(final_img_path, width=Cm(target_width_cm))
    except Exception as e:
        print(f"[警告] 图片尺寸计算失败，使用默认宽度插入: {e}")
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(final_img_path, width=Cm(target_width_cm))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_after = Pt(12)
    run_cap = p_cap.add_run(caption_text)
    set_font(run_cap, 10.5, is_bold=False, font_name_cn="仿宋")


def prompt_value(prompt, default=""):
    text = input(prompt).strip()
    return text or default


def extract_leading_number(text, default="1"):
    match = re.match(r"^\s*(\d+)", str(text))
    return match.group(1) if match else str(default)


def build_runtime_context(cfg):
    ctx = {
        "config": cfg,
        "project": copy.deepcopy(cfg.get("project", {})),
        "runtime_paths": {},
        "defaults": copy.deepcopy(cfg.get("defaults", {})),
        "templates": copy.deepcopy(cfg.get("templates", {})),
        "data": {},
        "assets_resolved": {},
        "user_inputs": {},
        "output_name": None,
        "output_path": None,
        "chapter_order": [],
        "chapter_numbers": {},
        "counters": {"figure": {}, "table": {}},
    }

    runtime_paths = copy.deepcopy(cfg.get("runtime_paths", {}))
    # 多轮解析 runtime_paths 中的占位符
    for _ in range(10):
        prev = copy.deepcopy(runtime_paths)
        temp_context = copy.deepcopy(ctx)
        temp_context["runtime_paths"] = runtime_paths
        runtime_paths = resolve_object(runtime_paths, temp_context)
        if prev == runtime_paths:
            break
    ctx["runtime_paths"] = runtime_paths

    # 预先登记 chapter id -> chapter number
    chapter_index = 0
    for block in cfg.get("report_structure", []):
        if block.get("type") == "chapter":
            chapter_index += 1
            block_id = block.get("id", f"chapter_{chapter_index}")
            title = block.get("title", str(chapter_index))
            chapter_no = extract_leading_number(title, default=str(chapter_index))
            ctx["chapter_order"].append(block_id)
            ctx["chapter_numbers"][block_id] = chapter_no
    return ctx


def collect_data_source(source_name, source_cfg, ctx):
    source_cfg = resolve_object(copy.deepcopy(source_cfg), ctx)
    stype = source_cfg.get("type")

    if stype == "mintpy_attribute":
        from mintpy.utils import readfile
        atr = readfile.read_attribute(source_cfg["path"])
        return parse_sar_metadata(atr)

    if stype == "text_file":
        path = source_cfg["path"]
        parser = source_cfg.get("parser", "inline_text")
        if parser == "date_list":
            return read_date_list_from_txt(path)
        if parser == "inline_text":
            return read_text_file_safely(path)
        raise ValueError(f"未知 text_file parser: {parser}")

    if stype == "cfg_file":
        return parse_cfg_file(source_cfg["path"])

    if stype == "glob":
        pattern = source_cfg["pattern"]
        files = glob.glob(pattern)
        sort_mode = source_cfg.get("sort", "default")
        if sort_mode == "natural":
            def natural_key(value):
                return [int(x) if x.isdigit() else x.lower() for x in re.split(r"(\d+)", value)]
            files = sorted(files, key=natural_key)
        else:
            files = sorted(files)
        return [{"path": p, "name": os.path.basename(p)} for p in files]

    if stype == "interactive":
        prompt = source_cfg.get("prompt", ">>> ")
        default = source_cfg.get("default", "")
        return prompt_value(prompt, default=default)

    if stype == "literal":
        return source_cfg.get("value")

    raise ValueError(f"未知数据源类型: {stype}")


def collect_all_data_sources(cfg, ctx):
    for name, source_cfg in cfg.get("data_sources", {}).items():
        ctx["data"][name] = collect_data_source(name, source_cfg, ctx)
    return ctx


def resolve_asset(asset_name, ctx):
    if asset_name in ctx["assets_resolved"]:
        return ctx["assets_resolved"][asset_name]

    if asset_name not in ctx["config"].get("assets", {}):
        raise KeyError(f"assets 中未定义: {asset_name}")

    asset_cfg = resolve_object(copy.deepcopy(ctx["config"]["assets"][asset_name]), ctx)
    atype = asset_cfg.get("type")

    if atype == "figure":
        chosen_path = None
        if "path" in asset_cfg:
            chosen_path = asset_cfg["path"]
        elif "candidates" in asset_cfg:
            for path in asset_cfg["candidates"]:
                if os.path.exists(path):
                    chosen_path = path
                    break
        result = {
            "type": "figure",
            "path": chosen_path,
            "exists": bool(chosen_path and os.path.exists(chosen_path)),
        }

    elif atype == "figure_group":
        source_name = asset_cfg.get("source")
        items = ctx["data"].get(source_name, [])
        result = {
            "type": "figure_group",
            "items": items,
            "exists": len(items) > 0,
        }
    else:
        raise ValueError(f"未知 asset 类型: {atype}")

    ctx["assets_resolved"][asset_name] = result
    return result


def asset_exists(asset_name, ctx):
    return resolve_asset(asset_name, ctx).get("exists", False)


def source_nonempty(source_name, ctx):
    value = ctx["data"].get(source_name)
    if value is None:
        return False
    if isinstance(value, (list, tuple, dict, str)):
        return len(value) > 0
    return bool(value)


def next_counter(counter_type, chapter_id, ctx):
    store = ctx["counters"].setdefault(counter_type, {})
    key = chapter_id or "__global__"
    store.setdefault(key, 0)
    store[key] += 1
    return store[key]


def make_figure_no(chapter_id, ctx):
    idx = next_counter("figure", chapter_id, ctx)
    chapter_no = ctx["chapter_numbers"].get(chapter_id, "1")
    return f"{chapter_no}-{idx}"


def make_table_no(chapter_id, ctx):
    idx = next_counter("table", chapter_id, ctx)
    chapter_no = ctx["chapter_numbers"].get(chapter_id, "1")
    return f"{chapter_no}-{idx}"


def build_render_context(ctx, current_chapter_id=None, loop_vars=None, extra=None):
    render_ctx = {
        "project": ctx.get("project", {}),
        "runtime_paths": ctx.get("runtime_paths", {}),
        "defaults": ctx.get("defaults", {}),
        "templates": ctx.get("templates", {}),
        "data": ctx.get("data", {}),
        "user_inputs": ctx.get("user_inputs", {}),
        "current_chapter_id": current_chapter_id,
        "current_chapter_no": ctx["chapter_numbers"].get(current_chapter_id, ""),
        "output_name": ctx.get("output_name"),
        "output_path": ctx.get("output_path"),
    }
    if loop_vars:
        render_ctx.update(loop_vars)
    if extra:
        render_ctx.update(extra)
    return render_ctx


def render_report(doc, cfg, ctx):
    for block in cfg.get("report_structure", []):
        render_block(doc, block, ctx, current_chapter_id=None, loop_vars=None)


def render_block(doc, block, ctx, current_chapter_id=None, loop_vars=None):
    render_ctx = build_render_context(ctx, current_chapter_id=current_chapter_id, loop_vars=loop_vars)
    block = resolve_object(copy.deepcopy(block), render_ctx)
    btype = block.get("type")

    if btype == "cover":
        return render_cover(doc, block, ctx)

    if btype == "toc":
        return render_toc(doc, block, ctx)

    if btype in ("chapter", "section", "subsection"):
        return render_heading_block(doc, block, ctx, current_chapter_id)

    if btype == "paragraph":
        return render_paragraph_block(doc, block, ctx)

    if btype == "software_versions":
        return render_software_versions_block(doc, block, ctx)

    if btype == "kv_line":
        return render_kv_line_block(doc, block, ctx)

    if btype == "figure":
        return render_figure_block(doc, block, ctx, current_chapter_id, loop_vars)

    if btype == "metadata_table":
        return render_metadata_table_block(doc, block, ctx)

    if btype == "table":
        return render_table_block(doc, block, ctx, current_chapter_id)

    if btype == "cfg_table":
        return render_cfg_table_block(doc, block, ctx)

    if btype == "text_file_line":
        return render_text_file_line_block(doc, block, ctx)

    if btype == "interactive_text":
        return render_interactive_text_block(doc, block, ctx)

    if btype == "conditional":
        return render_conditional_block(doc, block, ctx, current_chapter_id, loop_vars)

    if btype == "repeat":
        return render_repeat_block(doc, block, ctx, current_chapter_id)

    if btype == "page_break":
        doc.add_page_break()
        return None

    raise ValueError(f"未知 block 类型: {btype}")


def render_cover(doc, block, ctx):
    params = copy.deepcopy(block.get("params", {}))
    title_template = ctx.get("templates", {}).get("headings", {}).get(
        "cover_title",
        "基于{data_source}数据的{region_name}地区\n时序InSAR地表形变处理与分析报告",
    )

    cover_values = {}
    prompts = {
        "data_source": "1. 请输入数据源",
        "region_name": "2. 请输入地区",
        "lead_author": "3. 请输入第一作者",
    }

    print("请输入封面信息 ：")
    for key, spec in params.items():
        if key in ("logo_path", "output_name_template"):
            continue
        if isinstance(spec, str) and spec == "interactive":
            default_key = f"default_{key}"
            default_val = block.get(default_key, "")
            prompt_label = prompts.get(key, f"请输入 {key}")
            cover_values[key] = prompt_value(f"{prompt_label} [默认: {default_val}]: ", default=default_val)
        else:
            cover_values[key] = spec

    ctx["user_inputs"].update(cover_values)

    output_name_template = params.get("output_name_template", "{region_name}InSAR处理报告.docx")
    output_name = resolve_string_template(output_name_template, build_render_context(ctx, extra=cover_values))
    ctx["output_name"] = output_name
    output_dir = ctx["runtime_paths"].get("output_dir", os.getcwd())
    ctx["output_path"] = os.path.join(output_dir, output_name)

    section = doc.sections[0]
    section.top_margin = Cm(3.0)
    section.bottom_margin = Cm(2.5)

    p_logo = doc.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo_path = resolve_string_template(str(params.get("logo_path", "")), build_render_context(ctx))
    if logo_path and os.path.exists(logo_path):
        p_logo.add_run().add_picture(logo_path, width=Cm(5))
    elif logo_path:
        print(f"[提示] 未找到 Logo 图片: {logo_path}")
    p_logo.paragraph_format.space_after = Pt(130)

    title_text = resolve_string_template(title_template, build_render_context(ctx, extra=cover_values))
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.line_spacing = 1.5
    run_title = p_title.add_run(title_text)
    set_font(run_title, font_size_pt=24, is_bold=True, font_name_cn="仿宋")
    p_title.paragraph_format.space_after = Pt(150)

    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_ALIGN_PARAGRAPH.LEFT
    table.autofit = False
    col_width_label = Cm(4.5)
    col_width_value = Cm(9.5)

    for row in table.rows:
        row.cells[0].width = col_width_label
        row.cells[1].width = col_width_value
        tr_height = OxmlElement("w:trHeight")
        tr_height.set(qn("w:val"), "650")
        row._tr.get_or_add_trPr().append(tr_height)

    def fill_row(row_idx, label, value):
        cell_lbl = table.cell(row_idx, 0)
        p_l = cell_lbl.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_l = p_l.add_run(label)
        set_font(run_l, 18, is_bold=True, font_name_cn="仿宋")
        cell_lbl.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        cell_val = table.cell(row_idx, 1)
        p_v = cell_val.paragraphs[0]
        p_v.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_v = p_v.add_run(value)
        set_font(run_v, 18, is_bold=True, font_name_cn="仿宋")
        cell_val.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    fill_row(0, "单　　位：", "中国科学院空天信息创新研究院")
    fill_row(1, "作　　者：", f"{cover_values.get('lead_author', 'yourname')}、张云俊")
    fill_row(2, "日　　期：", datetime.now().strftime("%Y 年 %m 月 %d 日"))

    print("-" * 30)
    print("参数设置完成：")
    print(f"数据源: {cover_values.get('data_source', '')}")
    print(f"地区: {cover_values.get('region_name', '')}")
    print(f"第一作者: {cover_values.get('lead_author', '')}")
    print("-" * 30)

    doc.add_page_break()


def render_toc(doc, block, ctx):
    p_toc_title = doc.add_paragraph()
    p_toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_title.paragraph_format.space_before = Pt(20)
    p_toc_title.paragraph_format.space_after = Pt(20)

    run_toc_title = p_toc_title.add_run("目    录")
    set_font(run_toc_title, font_size_pt=16, is_bold=True, font_name_cn="仿宋")

    create_toc_element(doc)
    doc.add_page_break()


def render_heading_block(doc, block, ctx, current_chapter_id=None):
    btype = block["type"]
    level = block.get("level")
    if level is None:
        level = 1 if btype == "chapter" else 2 if btype == "section" else 3

    title = block.get("title", "")
    heading = doc.add_heading(title, level=level)
    run = heading.runs[0] if heading.runs else heading.add_run()
    font_size = 16 if level == 1 else 14 if level == 2 else 12
    set_font(run, font_size, True, font_name_cn="黑体")
    heading.paragraph_format.space_after = Pt(12 if level == 1 else 6)

    new_chapter_id = current_chapter_id
    if btype == "chapter":
        new_chapter_id = block.get("id", current_chapter_id)

    for child in block.get("children", []):
        render_block(doc, child, ctx, current_chapter_id=new_chapter_id)

    if btype == "chapter" and ctx["defaults"].get("auto_page_break_after_chapter", True):
        doc.add_page_break()


def render_paragraph_block(doc, block, ctx):
    text = block.get("text", "")
    p = doc.add_paragraph(text)
    p.paragraph_format.line_spacing = block.get("line_spacing", 1.5)
    if block.get("first_line_indent", False):
        p.paragraph_format.first_line_indent = Pt(24)
    run = p.runs[0] if p.runs else p.add_run()
    set_font(run, block.get("font_size_pt", 12), block.get("is_bold", False), block.get("font_name_cn", "仿宋"))




def render_software_versions_block(doc, block, ctx):
    """
    渲染“处理软件”信息行。
    支持 schema 中的 block:
      - type: software_versions
        label: 处理软件
        platform_source: mintpy_meta
        stack_processor: auto
    """
    label = block.get("label", "处理软件")
    platform_source = block.get("platform_source", "mintpy_meta")
    stack_processor = block.get("stack_processor", "auto")

    vers = get_software_versions()

    meta = ctx.get("data", {}).get(platform_source, {}) or {}

    if stack_processor == "auto":
        satellite = str(meta.get("satellite", "")).lower()
        if "sentinel" in satellite or "sen" in satellite:
            stack_proc = "topsStack"
        elif "alos" in satellite:
            stack_proc = "stripmapStack"
        else:
            stack_proc = "stripmapStack"
    else:
        stack_proc = str(stack_processor)

    soft_str = f"ISCE-2 ({vers['ISCE']}) / {stack_proc} + MintPy ({vers['MintPy']})"

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5

    r1 = p.add_run(f"{label}：")
    set_font(r1, 12, False, '仿宋')

    r2 = p.add_run(soft_str)
    set_font(r2, 12, False, 'Times New Roman')

    return None

def render_kv_line_block(doc, block, ctx):
    label = block.get("label", "")
    source_name = block.get("source")
    value = ctx["data"].get(source_name, "") if source_name else block.get("value", "")

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    r1 = p.add_run(f"{label}：")
    set_font(r1, 12, False, "仿宋")
    r2 = p.add_run(str(value))
    set_font(r2, 12, False, "Times New Roman")


def render_metadata_table_block(doc, block, ctx):
    source_name = block.get("source")
    meta = ctx["data"].get(source_name, {})

    extra_pairs = block.get("extra_pairs", [])
    fields = list(block.get("fields", []))
    for pair in extra_pairs:
        if isinstance(pair, list) and len(pair) == 2:
            fields.append(pair)

    for label, key in fields:
        value = meta.get(key, "N/A")
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.25
        run_lbl = p.add_run(f"{label}：")
        set_font(run_lbl, 12, False, "仿宋")
        run_val = p.add_run(str(value))
        set_font(run_val, 12, False, "Times New Roman")


def render_date_grid_table(doc, title, values, columns, current_chapter_id, ctx):
    table_no = make_table_no(current_chapter_id, ctx)
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(12)
    p_cap.paragraph_format.space_after = Pt(6)

    final_title = title
    if "{table_no}" in title:
        final_title = title.format(table_no=table_no)

    run_cap = p_cap.add_run(final_title)
    set_font(run_cap, 12, is_bold=False, font_name_cn="仿宋")

    num_cols = max(1, int(columns))
    num_rows = max(1, math.ceil(len(values) / num_cols))
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = ctx["defaults"].get("table_style", "Table Grid")
    table.autofit = True

    for i, value in enumerate(values):
        row_idx = i // num_cols
        col_idx = i % num_cols
        cell = table.cell(row_idx, col_idx)
        cell.text = str(value)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.25
        for run in p.runs:
            set_font(run, 12, is_bold=False)


def render_table_block(doc, block, ctx, current_chapter_id):
    source_name = block.get("source")
    renderer = block.get("renderer", "date_grid")
    data = ctx["data"].get(source_name, [])
    if not data:
        p = doc.add_paragraph("（未检测到表格数据）")
        run = p.runs[0] if p.runs else p.add_run()
        set_font(run, 10.5, False, "仿宋")
        return

    if renderer == "date_grid":
        render_date_grid_table(
            doc=doc,
            title=block.get("title", "表{table_no} 数据表"),
            values=data,
            columns=block.get("columns", 4),
            current_chapter_id=current_chapter_id,
            ctx=ctx,
        )
        return

    raise ValueError(f"未知 table renderer: {renderer}")


def render_cfg_table_block(doc, block, ctx):
    title = block.get("title", "参数配置")
    source_name = block.get("source")
    cfg_params = ctx["data"].get(source_name, [])
    filter_cfg = block.get("filter", {})
    exclude_values = {str(v).lower() for v in filter_cfg.get("exclude_values", [])}

    filtered = []
    for k, v in cfg_params:
        if str(v).lower() in exclude_values:
            continue
        filtered.append((k, v))

    p_mint = doc.add_paragraph()
    p_mint.paragraph_format.space_before = Pt(12)
    p_mint.paragraph_format.space_after = Pt(6)
    r_mint = p_mint.add_run(title)
    set_font(r_mint, 12, True, font_name_cn="黑体")

    if not filtered:
        p_none = doc.add_paragraph("（未检测到非默认配置）")
        set_font(p_none.runs[0] if p_none.runs else p_none.add_run(), 10.5, False, "仿宋")
        return

    table = doc.add_table(rows=len(filtered) + 1, cols=2)
    table.style = ctx["defaults"].get("table_style", "Table Grid")
    table.autofit = True

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "参数 (Parameter)"
    hdr_cells[1].text = "值 (Value)"
    for cell in hdr_cells:
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(p.runs[0] if p.runs else p.add_run(), 12, True, "黑体")

    for i, (key, val) in enumerate(filtered):
        row_cells = table.rows[i + 1].cells
        row_cells[0].text = str(key)
        row_cells[1].text = str(val)

        p_k = row_cells[0].paragraphs[0]
        p_k.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_font(p_k.runs[0] if p_k.runs else p_k.add_run(), 10.5, False, "Times New Roman")

        p_v = row_cells[1].paragraphs[0]
        p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(p_v.runs[0] if p_v.runs else p_v.add_run(), 10.5, False, "Times New Roman")


def render_text_file_line_block(doc, block, ctx):
    source_name = block.get("source")
    value = ctx["data"].get(source_name)
    if not value:
        value = "[未找到或无内容]"
    template = block.get("template", "{value}")
    text = resolve_string_template(template, build_render_context(ctx, extra={"value": value}))
    add_text_line(
        doc,
        text,
        font_size_pt=block.get("font_size_pt", 10.5),
        font_name_cn=block.get("font_name_cn", "仿宋"),
        space_after_pt=block.get("space_after_pt", 6),
    )


def render_interactive_text_block(doc, block, ctx):
    source_name = block.get("source")
    value = ctx["data"].get(source_name, "")
    if not value:
        value = block.get("default", "")
    p = doc.add_paragraph(value)
    p.paragraph_format.line_spacing = block.get("line_spacing", 1.5)
    p.paragraph_format.first_line_indent = Pt(block.get("first_line_indent_pt", 24))
    run = p.runs[0] if p.runs else p.add_run()
    set_font(run, block.get("font_size_pt", 12), False, block.get("font_name_cn", "仿宋"))


def render_figure_block(doc, block, ctx, current_chapter_id, loop_vars=None):
    if "asset" in block:
        asset = resolve_asset(block["asset"], ctx)
        path = asset.get("path")
    else:
        path = block.get("path")

    policy = block.get("on_missing", ctx["defaults"].get("missing_asset_policy", "warn"))
    if not path or not os.path.exists(path):
        msg = f"[图片缺失: {os.path.basename(path) if path else 'unknown'}]"
        if policy == "error":
            raise FileNotFoundError(msg)
        if policy == "warn":
            p = doc.add_paragraph(msg)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p.runs[0] if p.runs else p.add_run(), 10.5, False, "Times New Roman")
        return

    figure_no = make_figure_no(current_chapter_id, ctx)
    render_ctx = build_render_context(ctx, current_chapter_id=current_chapter_id, loop_vars=loop_vars, extra={"figure_no": figure_no})
    caption = resolve_string_template(block.get("caption", ""), render_ctx)

    width_cm = float(block.get("width_cm", ctx["defaults"].get("figure_width_cm", 14.0)))
    max_height_cm = float(block.get("max_height_cm", ctx["defaults"].get("figure_max_height_cm", 14.0)))
    pdf_render_dpi = int(block.get("pdf_render_dpi", ctx["defaults"].get("pdf_render_dpi", 300)))

    add_figure_with_caption(
        doc,
        path,
        caption,
        target_width_cm=width_cm,
        max_height_cm=max_height_cm,
        pdf_render_dpi=pdf_render_dpi,
    )


def render_conditional_block(doc, block, ctx, current_chapter_id, loop_vars=None):
    when = block.get("when", {})
    ok = False
    if "asset_exists" in when:
        ok = asset_exists(when["asset_exists"], ctx)
    elif "source_nonempty" in when:
        ok = source_nonempty(when["source_nonempty"], ctx)
    else:
        raise ValueError("conditional.when 目前仅支持 asset_exists/source_nonempty")

    branch = block.get("then", []) if ok else block.get("else", [])
    for child in branch:
        render_block(doc, child, ctx, current_chapter_id=current_chapter_id, loop_vars=loop_vars)


def render_repeat_block(doc, block, ctx, current_chapter_id):
    source_name = block.get("source")
    items = ctx["data"].get(source_name, [])
    alias = block.get("as", "item")
    for idx, item in enumerate(items, start=1):
        loop_vars = {
            alias: item,
            "item": item,
            "item_index": idx,
            "loop_index": idx,
        }
        for child in block.get("children", []):
            render_block(doc, child, ctx, current_chapter_id=current_chapter_id, loop_vars=loop_vars)


def build_report_from_yaml(doc, yaml_path=None):
    cfg = load_report_config(yaml_path=yaml_path)
    ctx = build_runtime_context(cfg)
    ctx = collect_all_data_sources(cfg, ctx)
    render_report(doc, cfg, ctx)
    return doc, ctx


def build_report_and_save(yaml_path=None):
    cfg = load_report_config(yaml_path=yaml_path)
    ctx = build_runtime_context(cfg)
    ctx = collect_all_data_sources(cfg, ctx)

    output_dir = ctx["runtime_paths"].get("output_dir", os.getcwd())
    os.makedirs(output_dir, exist_ok=True)

    doc = Document()
    render_report(doc, cfg, ctx)

    if not ctx.get("output_name"):
        ctx["output_name"] = "InSAR_Report.docx"
    if not ctx.get("output_path"):
        ctx["output_path"] = os.path.join(output_dir, ctx["output_name"])

    doc.save(ctx["output_path"])
    print("-" * 30)
    print(f"[成功] 文档已生成: {ctx['output_path']}")
    print("操作提示：请打开 Word 文档，在目录页【右键 -> 更新域 -> 更新整个目录】")
    print("-" * 30)
    return ctx["output_path"], ctx


if __name__ == "__main__":
    build_report_and_save()
