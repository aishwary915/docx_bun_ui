import { type RouteConfig, index, route } from "@react-router/dev/routes";

export default [
  index("routes/home.tsx"),
  route("doc_viewer", "routes/doc_viewer.tsx"),
  route("api/export", "routes/api.export.ts"),
] satisfies RouteConfig;
