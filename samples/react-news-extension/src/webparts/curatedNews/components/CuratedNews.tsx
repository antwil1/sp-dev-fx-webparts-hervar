import * as React from "react";
import styles from "./CuratedNews.module.scss";
import { ICuratedNewsProps } from "./ICuratedNewsProps";
import { Card, Col, Row, Space, Spin, Tag, Pagination } from "antd";
import Meta from "antd/lib/card/Meta";
import SPService from "../../../services/SPService";
import { ISearchResult } from "@pnp/sp/search";
import GraphService from "../../../services/GraphService";
import CachingService from "../../../services/CachingService";
import { ITerm } from "../../preferences/types/Component.Types";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

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
  const PAGE_SIZE = 2;

  const [data, setData] = React.useState<ISearchResult[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [page, setPage] = React.useState(1);
  const [total, setTotal] = React.useState(0);

  const preferenceCacheKey = `CuratedNews-UserPreferences-${loginName}`;

  const onConfigure = () => context.propertyPane.open();

  const getUserPreferences = React.useCallback(async () => {
    const cachedData = CachingService.get(preferenceCacheKey);
    if (cachedData !== null) return cachedData;

    const result = await GraphService.GetPreferences(extensionName);
    if (result && result.Tags && result.Tags.length > 0 && enableCaching) {
      CachingService.set(preferenceCacheKey, result.Tags);
    }
    return result?.Tags || [];
  }, [preferenceCacheKey, extensionName, enableCaching]);

  const fetchData = React.useCallback(async () => {
    setLoading(true);
    try {
      const tags = await getUserPreferences();
      if (!Array.isArray(tags) || tags.length === 0) {
        setData([]);
        setTotal(0);
        return;
      }

      const queryTemplate = composeQueryTemplate(tags);
      if (!queryTemplate) {
        setData([]);
        setTotal(0);
        return;
      }

      const { items, total } = await SPService.getSearchResults(
        queryTemplate,
        managedPropertyName,
        DISPLAY_PROP,
        page,
        PAGE_SIZE
      );

      setData(items ?? []);
      setTotal(total ?? 0);
    } catch (err) {
      console.error("fetchData error", err);
      setData([]);
      setTotal(0);
    } finally {
      setLoading(false);
    }
  }, [getUserPreferences, managedPropertyName, DISPLAY_PROP, page]);

  // initial load
  React.useEffect(() => {
    fetchData();
  }, [fetchData]);

  // lyssna på "preferencesSaved"
  React.useEffect(() => {
    const handler = (e: Event) => {
      const d = (e as CustomEvent).detail || {};
      if (d.loginName && d.loginName !== loginName) return;

      CachingService.remove(`CuratedNews-UserPreferences-${loginName}`);
      setPage(1); // reset to first page
    };
    window.addEventListener("curated:preferencesSaved", handler);
    return () => window.removeEventListener("curated:preferencesSaved", handler);
  }, [loginName]);

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
          <Card
            title={title}
            headStyle={{ fontSize: "2rem" }}
            extra={<a href={newsPageLink}>Visa alla</a>}
          >
            <Row gutter={16}>
              {data.length > 0 &&
                data.map((newsItem: any) => {
                  const raw: string | undefined =
                    newsItem[DISPLAY_PROP] ?? newsItem[managedPropertyName];

                  const tags: string[] = raw
                    ? raw
                        .split(";")
                        .map((s) => (s.includes("|") ? s.split("|")[0] : s))
                        .map((s) => s.trim())
                        .filter(Boolean)
                    : [];

                  return (
                    <Col key={newsItem.DocId} xs={24} md={6}>
                      <Card
                        className={styles.newsCard}
                        hoverable
                        bordered={false}
                        style={{ cursor: "pointer" }}
                        onClick={() => (window.location.href = newsItem.Path)}
                        cover={
                          <img
                            alt={newsItem.Title}
                            src={newsItem.PictureThumbnailURL}
                          />
                        }
                        actions={[
                          <div
                            key={`tags-${newsItem.DocId}`}
                            style={{ width: "100%" }}
                            onClick={(e) => e.stopPropagation()}
                          >
                            <Space size={[8, 8]} wrap className={styles.tags}>
                              {tags.map((tag) => (
                                <Tag key={tag} color="#EDEBE9">
                                  {tag}
                                </Tag>
                              ))}
                            </Space>
                          </div>,
                        ]}
                      >
                        <Meta
                          title={<a href={newsItem.Path}>{newsItem.Title}</a>}
                          description={
                            <>
                              <span className={styles.description}>
                                {newsItem.Description}
                              </span>
                              <div style={{ marginTop: 10 }} />
                            </>
                          }
                        />
                      </Card>
                    </Col>
                  );
                })}
            </Row>

            {total > PAGE_SIZE && (
              <div className={styles.pagination}>
                <Pagination
                  current={page}
                  pageSize={PAGE_SIZE}
                  total={total}
                  showSizeChanger={false}
                  onChange={(p) => setPage(p)}
                />
              </div>
            )}
          </Card>
        </Spin>
      </div>
    </section>
  );

  function composeQueryTemplate(tags: ITerm[]) {
    if (!Array.isArray(tags) || tags.length === 0) return null;

    const taxValues = `(${tags.map((t) => t.id).join(" OR ")})`;
    const filter = `({|${managedPropertyName}:${taxValues}})`;

    if (customQueryTemplate && customQueryTemplate.trim().length > 0) {
      const tpl = customQueryTemplate.trim();
      return tpl.includes("{FILTER}")
        ? tpl.replace("{FILTER}", filter)
        : `${tpl} ${filter}`;
    }

    return `{searchTerms} (ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118*) PromotedState=2 ${filter}`;
  }
};
