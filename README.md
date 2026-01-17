# OneBackoffice

통합 회사·재고·쇼핑몰 관리 SaaS (MVP 목표)입니다.

## Monorepo 구조
```
packages/
  backend/   # NestJS API
  web/       # Next.js Admin
  shared/    # 공용 타입/유틸
```

## 로컬 개발
```bash
pnpm install
pnpm dev
```

## 인프라 의존성
```bash
docker-compose up -d
```

## 주요 방향
- 모듈러 모놀리식 (backend/web/shared)
- integrations는 plugin 인터페이스 기반
- 재고는 append-only ledger로 관리
