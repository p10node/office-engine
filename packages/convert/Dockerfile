# office-convert as an HTTP service. Built from the repository root:
#   docker build -f packages/convert/Dockerfile -t office-convert .
#   docker run --rm -p 3000:3000 office-convert
# Settings: PORT, HOST, WORKERS (conversions at once), TIMEOUT (seconds per file), MAX_SIZE (MB per request), MAX_FILES.

FROM oven/bun:1.4 AS build
WORKDIR /src
COPY package.json bun.lock bunfig.toml tsconfig.json LICENSE ./
RUN bun install --frozen-lockfile
COPY src src
COPY packages packages
COPY scripts/build-packages.ts scripts/
RUN bun scripts/build-packages.ts convert

# the package has no dependency: the runtime image is Node and the bundle
FROM node:26-slim
WORKDIR /app
COPY --from=build /src/packages/convert/package.json ./
COPY --from=build /src/packages/convert/dist ./dist
ENV NODE_ENV=production PORT=3000
USER node
EXPOSE 3000
HEALTHCHECK --interval=30s --timeout=5s CMD ["node", "-e", "fetch(`http://127.0.0.1:${process.env.PORT}/health`).then((r) => process.exit(r.ok ? 0 : 1), () => process.exit(1))"]
ENTRYPOINT ["node", "dist/cli.js"]
CMD ["serve"]
