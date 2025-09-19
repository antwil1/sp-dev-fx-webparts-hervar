import * as React from "react";
import styles from "./CuratedNews.module.scss";
import { ICuratedNewsProps } from "./ICuratedNewsProps";
import { Card, Col, Row, Space, Spin, Tag, Pagination } from "antd";
import Meta from "antd/lib/card/Meta";
import SPService from "../../../services/SPService";
import { ISearchResult } from "@pnp/sp/search";
import GraphService from "../../../services/GraphService";
import CachingService from "../../../services/CachingService";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

/** Responsiv page size */
function useResponsivePageSize() {
  const calc = React.useCallback(() => (window.innerWidth < 768 ? 1 : 4), []);
  const [pageSize, setPageSize] = React.useState<number>(calc);
  React.useEffect(() => {
    let raf: number | null = null;
    const onResize = () => {
      if (raf !== null) cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => setPageSize(calc()));
    };
    window.addEventListener("resize", onResize);
    return () => {
      if (raf !== null) cancelAnimationFrame(raf);
      window.removeEventListener("resize", onResize);
    };
  }, [calc]);
  return pageSize;
}

export const CuratedNews: React.FC<ICuratedNewsProps> = (props) => {
  const {
    extensionName,
    loginName,
    title,
    managedPropertyName,
    context,
    newsPageLink,
    enableCaching,
    customQueryTemplate,
  } = props;

  const DISPLAY_PROP = "RefinableString01";
  const responsivePageSize = useResponsivePageSize();

  const [data, setData] = React.useState<ISearchResult[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [page, setPage] = React.useState(1);
  const [total, setTotal] = React.useState(0);

  // Enkelt swipe-stöd
  const SWIPE_THRESHOLD = 48;
  const MAX_VERTICAL_DRIFT = 40;
  const touchStart = React.useRef<{ x: number; y: number } | null>(null);
  const onTouchStart = (e: React.TouchEvent) => {
    const t = e.changedTouches[0];
    touchStart.current = { x: t.clientX, y: t.clientY };
  };
  const onTouchEnd = (e: React.TouchEvent) => {
    if (!touchStart.current || loading) return;
    const t0 = touchStart.current;
    const t = e.changedTouches[0];
    const dx = t.clientX - t0.x;
    const dy = Math.abs(t.clientY - t0.y);
    if (dy > MAX_VERTICAL_DRIFT) { touchStart.current = null; return; }
    if (dx <= -SWIPE_THRESHOLD) setPage((p) => (p * responsivePageSize < total ? p + 1 : p));
    else if (dx >= SWIPE_THRESHOLD) setPage((p) => (p > 1 ? p - 1 : p));
    touchStart.current = null;
  };

  const preferenceCacheKey = `CuratedNews-UserPreferences-${loginName}`;
  const onConfigure = () => context.propertyPane.open();

  const getUserPreferences = React.useCallback(async (): Promise<string[]> => {
    const cachedData = CachingService.get<string[]>(preferenceCacheKey);
    if (cachedData !== null) return cachedData;

    const result = await GraphService.GetPreferences(extensionName);
    const ids: string[] = (result && Array.isArray(result.Tags)) ? result.Tags : [];
    if (ids.length > 0 && enableCaching) CachingService.set(preferenceCacheKey, ids);
    return ids;
  }, [preferenceCacheKey, extensionName, enableCaching]);

  const fetchData = React.useCallback(async () => {
    setLoading(true);
    try {
      const ids = await getUserPreferences(); // <-- string[]
      if (!Array.isArray(ids) || ids.length === 0) {
        setData([]); setTotal(0); return;
      }

      const queryTemplate = composeQueryTemplate(ids);
      if (!queryTemplate) { setData([]); setTotal(0); return; }

      const { items, total } = await SPService.getSearchResults(
        queryTemplate,
        managedPropertyName,
        DISPLAY_PROP,
        page,
        responsivePageSize
      );

      setData(items ?? []);
      setTotal(total ?? 0);
    } catch (err) {
      console.error("fetchData error", err);
      setData([]); setTotal(0);
    } finally {
      setLoading(false);
    }
  }, [getUserPreferences, managedPropertyName, DISPLAY_PROP, page, responsivePageSize]);

  React.useEffect(() => { fetchData(); }, [fetchData]);
  React.useEffect(() => { setPage(1); }, [responsivePageSize]);

  React.useEffect(() => {
    const handler = (e: Event) => {
      const d = (e as CustomEvent).detail || {};
      if (d.loginName && d.loginName !== loginName) return;

      CachingService.remove(`CuratedNews-UserPreferences-${loginName}`);
      if (page === 1) { fetchData(); } else { setPage(1); }
    };
    window.addEventListener("curated:preferencesSaved", handler);
    return () => window.removeEventListener("curated:preferencesSaved", handler);
  }, [loginName, page, fetchData]);

  if (!extensionName || !managedPropertyName || !newsPageLink) {
    return (
      <Placeholder
        iconName="Edit"
        iconText="Configure your web part"
        description="Please provide the Microsoft Graph open extension name and managed property name."
        buttonLabel="Configure"
        onConfigure={onConfigure}
      />
    );
  }

  return (
    <section>
      <div className={styles["news-container"]}>
        <Spin spinning={loading} tip="Loading...">
          <Card title={<h2 className={styles.sectionTitle}>{title}</h2>} headStyle={{}}>
            <div onTouchStart={onTouchStart} onTouchEnd={onTouchEnd} style={{ touchAction: "pan-y" }}>
              <Row gutter={16}>
                {data.length > 0 && data.map((newsItem: any) => {
                  const raw: string | undefined =
                    newsItem[DISPLAY_PROP] ?? newsItem[managedPropertyName];
                  const tags: string[] = raw
                    ? raw.split(";").map((s) => (s.includes("|") ? s.split("|")[0] : s)).map((s) => s.trim()).filter(Boolean)
                    : [];
                  return (
                    <Col key={newsItem.DocId} xs={24} lg={6}>
                      <Card
                        className={styles.newsCard}
                        hoverable
                        bordered={false}
                        style={{ cursor: "pointer" }}
                        onClick={() => (window.location.href = newsItem.Path)}
                        cover={<img alt={newsItem.Title} src={newsItem.PictureThumbnailURL} />}
                        actions={[
                          <div key={`tags-${newsItem.DocId}`} style={{ width: "100%" }} onClick={(e) => e.stopPropagation()}>
                            <Space size={[8, 8]} wrap className={styles.tags}>
                              {tags.map((tag) => (<Tag key={tag} color="#EDEBE9">{tag}</Tag>))}
                            </Space>
                          </div>,
                        ]}
                      >
                        <Meta
                          title={<a href={newsItem.Path}>{newsItem.Title}</a>}
                          description={<><span className={styles.description}>{newsItem.Description}</span><div style={{ marginTop: 10 }} /></>}
                        />
                      </Card>
                    </Col>
                  );
                })}
              </Row>

              {total > responsivePageSize && (
                <div className={styles.pagination}>
                  <Pagination
                    current={page}
                    pageSize={responsivePageSize}
                    total={total}
                    showSizeChanger={false}
                    onChange={(p) => setPage(p)}
                    simple={true}
                    showLessItems={true}
                  />
                </div>
              )}
            </div>
          </Card>
        </Spin>
      </div>
    </section>
  );

  /** Nu tar funktionen string[] (id:n) */
  function composeQueryTemplate(ids: string[]) {
    if (!Array.isArray(ids) || ids.length === 0) return null;

    const taxValues = `(${ids.join(" OR ")})`;
    const filter = `({|${managedPropertyName}:${taxValues}})`;

    if (customQueryTemplate && customQueryTemplate.trim().length > 0) {
      const tpl = customQueryTemplate.trim();
      return tpl.includes("{FILTER}") ? tpl.replace("{FILTER}", filter) : `${tpl} ${filter}`;
    }

    return `{searchTerms} (ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118*) PromotedState=2 ${filter}`;
  }
};
