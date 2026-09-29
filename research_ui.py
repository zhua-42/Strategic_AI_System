"""Streamlit adapter; reusable engine/export modules have no Streamlit dependency."""
import hashlib
import json
import re
import time


def render(client,target,kind,report_type,period,submitted=False,purpose="综合分析"):
    import streamlit as st
    import pandas as pd
    import plotly.graph_objects as go
    import research_agent as agent
    import research_documents as documents
    import product_features as pf
    state=st.session_state.setdefault("autonomous_research_state",{})
    st.caption("自动研究：检索原文 → 引文核对 → 预测三表 → 逐章成稿 → 导出检查。长报告至少20个内容章节、10张图表；资料不足会显示草稿状态。")
    if state.get("chapters") and "autonomous_package" not in st.session_state:
        st.info(f"已保留 {len(state['chapters'])} 个章节。使用同一研究条件再次点击生成，可接着完成。")
    if submitted:
        if not target.strip():
            st.warning("请输入公司或行业名称。")
        else:
            st.session_state.pop("autonomous_package",None)
            st.session_state.pop("autonomous_exports",None)
            bar=st.progress(0,text="开始自动研究")
            started=time.monotonic()
            before=dict(st.session_state.get("usage_stats",{}))
            def progress(value,text): bar.progress(min(1.,max(0.,value)),text=text)
            try:
                package=agent.generate(client,target.strip(),kind,report_type,period,progress,
                                      st.session_state.get("uploaded_report_text", ""),state,purpose=purpose)
                st.session_state["autonomous_package"]=package
                bar.progress(1.,text="研究完成")
                # Preserve existing history/favorites representation plus the structured package.
                markdown="\n\n".join("## "+c["title"]+"\n\n"+"\n\n".join(c["paragraphs"]) for c in package["chapters"])
                st.session_state.setdefault("history",[]).insert(0,{"query":target,"content":markdown,"data":{"dossier":package}})
                after=st.session_state.get("usage_stats",{})
                pf.record_usage(target,time.monotonic()-started,after.get("input_tokens",0)-before.get("input_tokens",0),
                                after.get("output_tokens",0)-before.get("output_tokens",0),[])
            except Exception as exc:
                st.error(str(exc))
                st.caption("同一条件再次生成会继续已完成的章节；更换标的或报告周期会创建新研究。")
    package=st.session_state.get("autonomous_package")
    if not package:
        return
    st.subheader(package["target"]+" · "+package["report_type"])
    st.info(package["status"])
    if st.button("收藏报告并生成分享链接",key="dossier_favorite"):
        markdown="\n\n".join("## "+c["title"]+"\n\n"+"\n\n".join(c["paragraphs"]) for c in package["chapters"])
        favorite=pf.add_favorite(package["target"],markdown,{"dossier":package})
        st.session_state["dossier_share_link"]=f"{pf.APP_URL}?report={favorite['id']}"
        st.success("已收藏；分享链接会展示当前报告与来源。")
    if st.session_state.get("dossier_share_link"):
        st.code(st.session_state["dossier_share_link"],language=None)
        st.caption("收藏位于当前部署实例，服务器重建后可能需要重新收藏。")
    st.caption(f"{len(package['chapters'])} 个内容章节 / {len(package['chapters'])} 张图表；检索时间 {package['created_at'][:19]}。最新财务期以引文为准。")
    report_tab,model_tab,evidence_tab,export_tab=st.tabs(["图文报告","自动 Pro Forma","来源与质量","下载报告"])
    with report_tab:
        for index,c in enumerate(package["chapters"],1):
            with st.expander(f"{index:02d} {c['title']}",expanded=index==1):
                st.markdown("**"+c["takeaway"]+"**")
                split=max(2,len(c["paragraphs"])//2)
                for text in c["paragraphs"][:split]: st.markdown(text)
                v=c["visual"]
                st.markdown(f"**图表 {index}：{v['title']}**")
                if v["kind"]=="table":
                    st.dataframe(pd.DataFrame(v["rows"],columns=v["headers"]),use_container_width=True,hide_index=True)
                elif v["kind"] in ("flow","timeline"):
                    st.image(documents.visual_png(v),use_container_width=True)
                else:
                    trace=go.Bar(x=v["labels"],y=v["values"]) if v["kind"]=="bar" else go.Scatter(x=v["labels"],y=v["values"],mode="lines+markers")
                    fig=go.Figure(trace)
                    fig.update_layout(height=330,yaxis_title=v.get("unit",""),font_family="LXGW WenKai, KaiTi, serif",margin=dict(t=20,b=70))
                    st.plotly_chart(fig,use_container_width=True,key=f"dossier_{index}")
                st.caption(v["source"])
                for text in c["paragraphs"][split:]: st.markdown(text)
    with model_tab:
        model=package["model"]
        st.write("基期："+model["base_period"]+"；单位："+model["unit"])
        st.caption(model["opening_note"])
        st.warning(model["note"])
        st.dataframe(pd.DataFrame([{ "参数":k,"取值":v,"依据":model["assumption_basis"][k]} for k,v in model["assumptions"].items()]),hide_index=True,use_container_width=True)
        st.caption("来源："+model["base_source"]+"；模型自动从核实数据推导可用比率，其余逐项标为情景假设。Excel 中可修改并重算。")
        for idx in [15,16,17,18,22]:
            visual=agent.model_visual(idx,model)
            st.write(agent.COMPANY_CHAPTERS[idx])
            st.dataframe(pd.DataFrame(visual["rows"],columns=visual["headers"]),hide_index=True,use_container_width=True)
            st.caption(visual["source"])
    with evidence_tab:
        for label,passed in package["checks"].items(): st.write(("✓ " if passed else "待完善：")+label)
        for doc in package["sources"]:
            with st.expander(f"[{doc['id']}] {doc['title']}"):
                if doc["url"].startswith("http"): st.link_button("查看原文",doc["url"])
                st.caption(f"读取方式：{doc['method']}；抓取时间：{doc['retrieved_at']}")
                st.write(doc["text"][:5000])
        st.dataframe(pd.DataFrame(package["facts"]),use_container_width=True)
        with st.expander("检索失败与引文拒绝日志"):
            st.json({"retrieval":package["audit"],"rejected_quotes":package["rejected"]})
        if st.button("清除本次缓存并重新取数",key="refresh_research"):
            for key in ("autonomous_research_state","autonomous_package","autonomous_exports"): st.session_state.pop(key,None)
            st.rerun()
    with export_tab:
        st.caption("PDF 内嵌楷体；PPT 为与 PDF 同版式的楷体保真版。Word 和 Excel 可编辑，查看时建议安装下方开源字体。各格式独立生成，失败不影响其他格式。")
        name=re.sub(r'[\\/:*?"<>|]',"_",package["target"])+"_"+package["report_type"]
        cache=st.session_state.setdefault("autonomous_exports",{})
        for ext,label,mime in [("pdf","PDF 完整图文报告","application/pdf"),
                               ("pptx","PPT 楷体保真版","application/vnd.openxmlformats-officedocument.presentationml.presentation"),
                               ("docx","Word 可编辑报告","application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                               ("xlsx","Excel 三表模型与证据","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]:
            if st.button("生成 "+label,key="build_dossier_"+ext):
                with st.spinner("生成并检查 "+label):
                    try:
                        if ext in ("pdf","pptx") and "pdf" not in cache: cache["pdf"]=documents.pdf_bytes(package)
                        if ext=="pptx": cache[ext]=documents.pptx_bytes(package,cache["pdf"])
                        if ext=="docx": cache[ext]=documents.docx_bytes(package)
                        if ext=="xlsx": cache[ext]=documents.evidence_bytes(package)
                    except Exception as exc: st.error(f"{label} 生成失败：{exc}")
            if ext in cache:
                st.download_button("下载 "+label,cache[ext],name+"."+ext,mime,key="download_dossier_"+ext)
        if documents.FONT_PATH.exists():
            st.download_button("下载开源楷体（霞鹜文楷）",documents.FONT_PATH.read_bytes(),"LXGWWenKai-Regular.ttf","font/ttf")
            st.caption("字体采用 OFL 1.1 授权；网站 assets/fonts/OFL-WenKai.txt 随部署保存授权全文。")
