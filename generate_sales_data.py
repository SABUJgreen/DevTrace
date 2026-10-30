"""
Generate complete Sales Prospecting Tracker & Analysis Dashboard for DevTrace (devtrace.io)
Includes:
- Excel Workbook (Sales_Prospecting_Dashboard.xlsx) with 5 styled sheets, formulas, and charts
- CSV exports for Notion & Google Sheets import
- Interactive Web Dashboard (index.html)
"""

import os
import csv
import json
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart, Series

# -------------------------------------------------------------
# 1. RAW DATA DEFINITIONS: 60 REALISTIC TARGET COMPANIES (ICP)
# -------------------------------------------------------------
TARGET_COMPANIES = [
    {
        "id": "LEAD-001",
        "company": "Hasura",
        "contact_name": "Tanmai Gopal",
        "title": "VP of Engineering & Co-Founder",
        "email": "tanmai@hasura.io",
        "size": "250-500",
        "funding": "Series C ($100M+)",
        "tech_stack": "Go, Haskell, Node.js, GraphQL, PostgreSQL, Kubernetes",
        "personalization": "Noticed Hasura 3.0 v3 engine rebuild in Go; mocking complex multi-source upstream databases during CI likely consumes 20%+ of dev sprint time.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 95
    },
    {
        "id": "LEAD-002",
        "company": "Supabase",
        "contact_name": "Paul Copplestone",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "paul@supabase.com",
        "size": "100-250",
        "funding": "Series B ($116M)",
        "tech_stack": "Elixir, Go, TypeScript, PostgreSQL, Docker, Kubernetes",
        "personalization": "Saw GitHub release on Supabase Auth & Realtime engine; high rate of open-source PRs means automated regression test generation from live traffic would save hundreds of triage hours.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 98
    },
    {
        "id": "LEAD-003",
        "company": "Doppler",
        "contact_name": "Brian Vallelunga",
        "title": "VP of Engineering",
        "email": "brian@doppler.com",
        "size": "50-100",
        "funding": "Series A ($20M)",
        "tech_stack": "Go, Python, React, AWS, Docker, Microservices",
        "personalization": "Secret Ops platform handling millions of daily sync calls; zero-downtime microservices testing with DevTrace's zero-code eBPF capture eliminates fragile mock servers.",
        "ab_variant": "Email A (Velocity)",
        "status": "Replied - Interested",
        "meetings_booked": 0,
        "lead_score": 88
    },
    {
        "id": "LEAD-004",
        "company": "Temporal Technologies",
        "contact_name": "Maxim Fateev",
        "title": "Head of Engineering & Co-Founder",
        "email": "maxim@temporal.io",
        "size": "200-500",
        "funding": "Series B ($100M+)",
        "tech_stack": "Go, Java, Python, TypeScript, gRPC, Distributed Systems",
        "personalization": "Building durable workflow engines; testing multi-step distributed saga failures requires hours of mock state building. DevTrace creates deterministic replays automatically.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 96
    },
    {
        "id": "LEAD-005",
        "company": "Resend",
        "contact_name": "Zeno Rocha",
        "title": "VP of Product Engineering",
        "email": "zeno@resend.com",
        "size": "20-50",
        "funding": "Seed / Series A ($15M)",
        "tech_stack": "Node.js, TypeScript, Next.js, Redis, PostgreSQL, AWS",
        "personalization": "API-first transactional email delivery; API contracts must stay 100% backwards compatible. Capturing production payloads verifies edge-case regressions instantly.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 92
    },
    {
        "id": "LEAD-006",
        "company": "Railway",
        "contact_name": "Jake Cooper",
        "title": "Director of Infrastructure & Eng",
        "email": "jake@railway.app",
        "size": "50-100",
        "funding": "Series A ($20M)",
        "tech_stack": "Rust, Go, TypeScript, Docker, Kubernetes, Linux",
        "personalization": "PaaS deployment platform managing millions of ephemeral containers; eBPF-based traffic recording catches networking regressions without injecting intrusive SDKs.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 5)",
        "meetings_booked": 0,
        "lead_score": 85
    },
    {
        "id": "LEAD-007",
        "company": "ClickHouse Cloud",
        "contact_name": "Alexey Milovidov",
        "title": "VP of Core Engineering",
        "email": "alexey@clickhouse.com",
        "size": "250-500",
        "funding": "Series B ($250M)",
        "tech_stack": "C++, Go, Python, Kubernetes, AWS, High-Throughput OLAP",
        "personalization": "Rapid release cycles on ClickHouse Cloud control plane in Go; automated test case generation saves devs from maintaining complex distributed mock clusters.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Replied - Objection",
        "meetings_booked": 0,
        "lead_score": 82
    },
    {
        "id": "LEAD-008",
        "company": "Retool",
        "contact_name": "David Hsu",
        "title": "Head of Core Platform Engineering",
        "email": "david@retool.com",
        "size": "500-1000",
        "funding": "Series C ($135M+)",
        "tech_stack": "Node.js, TypeScript, Go, PostgreSQL, Redis, Kubernetes",
        "personalization": "Retool integrates with 50+ third-party APIs (Stripe, Salesforce, etc.). Developers lose hours maintaining external mock responses; DevTrace auto-generates mocks in YAML.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 94
    },
    {
        "id": "LEAD-009",
        "company": "Inngest",
        "contact_name": "Tony Holdstock-Brown",
        "title": "VP of Engineering & Co-Founder",
        "email": "tony@inngest.com",
        "size": "20-50",
        "funding": "Series A ($10M)",
        "tech_stack": "Go, TypeScript, Kafka, PostgreSQL, Docker, AWS",
        "personalization": "Event-driven workflow execution engine; capturing live asynchronous webhook events into reproducible test stubs eliminates flaky async integration tests.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 7)",
        "meetings_booked": 0,
        "lead_score": 84
    },
    {
        "id": "LEAD-010",
        "company": "LaunchDarkly",
        "contact_name": "John Kodumal",
        "title": "VP of Engineering & Co-Founder",
        "email": "john@launchdarkly.com",
        "size": "500-1000",
        "funding": "Series D ($330M+)",
        "tech_stack": "Go, Java, Node.js, Kafka, Kubernetes, AWS",
        "personalization": "Evaluating feature flag evaluations at massive scale. DevTrace deduplicates high-volume microservices traffic to create minimal, non-redundant test suites.",
        "ab_variant": "Email A (Velocity)",
        "status": "Replied - Interested",
        "meetings_booked": 0,
        "lead_score": 89
    },
    {
        "id": "LEAD-011",
        "company": "Courier",
        "contact_name": "Troy Goode",
        "title": "VP of Platform Engineering",
        "email": "troy@courier.com",
        "size": "50-100",
        "funding": "Series A ($12M)",
        "tech_stack": "Node.js, Python, TypeScript, DynamoDB, Redis, AWS",
        "personalization": "Multi-channel notification routing with dozens of downstream SMS/email providers. Auto-generating mocks directly from network calls prevents staging environment drift.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 3)",
        "meetings_booked": 0,
        "lead_score": 79
    },
    {
        "id": "LEAD-012",
        "company": "Postman",
        "contact_name": "Ankit Sobti",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "ankit@postman.com",
        "size": "500-1000",
        "funding": "Series D ($430M+)",
        "tech_stack": "Node.js, Go, Python, React, PostgreSQL, Redis, AWS",
        "personalization": "Postman sets the standard for API tooling; DevTrace complements collection runs by generating the underlying integration tests and database stubs automatically.",
        "ab_variant": "Email A (Velocity)",
        "status": "Replied - Objection",
        "meetings_booked": 0,
        "lead_score": 91
    },
    {
        "id": "LEAD-013",
        "company": "Teleport",
        "contact_name": "Ev Kontsevoy",
        "title": "VP of Core Infrastructure",
        "email": "ev@goteleport.com",
        "size": "250-500",
        "funding": "Series C ($110M)",
        "tech_stack": "Go, Rust, Kubernetes, eBPF, Linux, gRPC",
        "personalization": "As an infrastructure access platform built in Go and eBPF, Teleport's engineering team will appreciate DevTrace's kernel-level socket capture that requires zero SDK changes.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 97
    },
    {
        "id": "LEAD-014",
        "company": "Grafana Labs",
        "contact_name": "Anthony Woods",
        "title": "VP of Engineering & Co-Founder",
        "email": "anthony@grafana.com",
        "size": "1000+",
        "funding": "Series D ($500M+)",
        "tech_stack": "Go, TypeScript, MySQL, Kafka, Prometheus, Kubernetes",
        "personalization": "Maintaining hundreds of data source plugins in Go/TypeScript; recording API calls and caching mock data reduces plugin integration test runtimes by 60%.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 10)",
        "meetings_booked": 0,
        "lead_score": 87
    },
    {
        "id": "LEAD-015",
        "company": "Sentry",
        "contact_name": "David Cramer",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "david@sentry.io",
        "size": "250-500",
        "funding": "Series E ($215M)",
        "tech_stack": "Python, Django, Go, Rust, ClickHouse, Kafka",
        "personalization": "Processing billions of telemetry events daily; testing ingest pipeline regressions against real payload edge-cases without staging environment overhead.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 95
    },
    {
        "id": "LEAD-016",
        "company": "Prisma",
        "contact_name": "Johannes Schickling",
        "title": "VP of Engineering & Co-Founder",
        "email": "johannes@prisma.io",
        "size": "50-100",
        "funding": "Series B ($40M)",
        "tech_stack": "TypeScript, Rust, PostgreSQL, MySQL, Docker",
        "personalization": "Next-gen ORM testing multi-database query generation. DevTrace intercepts database wire protocols (Postgres, Mongo) to generate accurate mock response files.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Replied - Interested",
        "meetings_booked": 0,
        "lead_score": 86
    },
    {
        "id": "LEAD-017",
        "company": "Checkly",
        "contact_name": "Hanno Cserhati",
        "title": "VP of Product Engineering",
        "email": "hanno@checklyhq.com",
        "size": "20-50",
        "funding": "Series A ($12M)",
        "tech_stack": "Node.js, TypeScript, Playwright, AWS, Serverless",
        "personalization": "Checkly champions synthetic monitoring; pairing synthetic endpoint checks with DevTrace's auto-generated backend regression mocks creates full-lifecycle assurance.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 3)",
        "meetings_booked": 0,
        "lead_score": 80
    },
    {
        "id": "LEAD-018",
        "company": "Linear",
        "contact_name": "Karri Saarinen",
        "title": "VP of Engineering & Co-Founder",
        "email": "karri@linear.app",
        "size": "50-100",
        "funding": "Series B ($35M)",
        "tech_stack": "TypeScript, GraphQL, Node.js, PostgreSQL, Redis",
        "personalization": "Linear's iconic ultra-fast sync engine relies heavily on optimistic offline state and GraphQL mutations. DevTrace auto-captures complex sequential API states.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 93
    },
    {
        "id": "LEAD-019",
        "company": "Brex",
        "contact_name": "Cosmin Nicolaescu",
        "title": "VP of Engineering",
        "email": "cosmin@brex.com",
        "size": "1000+",
        "funding": "Series D ($1.5B)",
        "tech_stack": "Elixir, Go, Kotlin, PostgreSQL, Kafka, Kubernetes",
        "personalization": "Fintech card ledger and payment settlements require zero tolerance for regression bugs. DevTrace creates regression suites from staging traffic without manual mock coding.",
        "ab_variant": "Email A (Velocity)",
        "status": "Replied - Objection",
        "meetings_booked": 0,
        "lead_score": 90
    },
    {
        "id": "LEAD-020",
        "company": "Ramp",
        "contact_name": "Karim Atiyeh",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "karim@ramp.com",
        "size": "500-1000",
        "funding": "Series D ($1.4B)",
        "tech_stack": "Python, FastAPI, React, PostgreSQL, AWS, Microservices",
        "personalization": "Ramp's explosive engineering growth means onboarding dozens of devs monthly. DevTrace helps new devs refactor Python FastAPI microservices with immediate test safety nets.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 96
    },
    {
        "id": "LEAD-021",
        "company": "Mercury",
        "contact_name": "Max Tagher",
        "title": "VP of Engineering & Co-Founder",
        "email": "max@mercury.com",
        "size": "250-500",
        "funding": "Series B ($150M+)",
        "tech_stack": "Haskell, TypeScript, Python, PostgreSQL, AWS",
        "personalization": "High-assurance banking backend; automated test generation with built-in data masking ensures PCI/PII compliance while achieving 90%+ code coverage.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 7)",
        "meetings_booked": 0,
        "lead_score": 88
    },
    {
        "id": "LEAD-022",
        "company": "Plaid",
        "contact_name": "Jean-Denis Greze",
        "title": "Head of Engineering",
        "email": "jd@plaid.com",
        "size": "1000+",
        "funding": "Series D ($700M+)",
        "tech_stack": "Go, Python, TypeScript, PostgreSQL, AWS, Microservices",
        "personalization": "Connecting to 12,000+ financial institutions means constant upstream API changes. DevTrace's AI test gap analysis isolates upstream schema breakages automatically.",
        "ab_variant": "Email A (Velocity)",
        "status": "Replied - Objection",
        "meetings_booked": 0,
        "lead_score": 92
    },
    {
        "id": "LEAD-023",
        "company": "Modern Treasury",
        "contact_name": "Sam Aarons",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "sam@moderntreasury.com",
        "size": "100-250",
        "funding": "Series C ($180M)",
        "tech_stack": "Ruby, Go, PostgreSQL, AWS, Microservices",
        "personalization": "Payment operations platform handling wire and ACH routing; automating complex banking API stubs allows continuous delivery without maintaining static fixture files.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 94
    },
    {
        "id": "LEAD-024",
        "company": "Unit.co",
        "contact_name": "Doron Somech",
        "title": "VP of Engineering & Co-Founder",
        "email": "doron@unit.co",
        "size": "100-250",
        "funding": "Series C ($170M)",
        "tech_stack": "Go, Node.js, TypeScript, PostgreSQL, Docker, AWS",
        "personalization": "Banking-as-a-service requires testing multi-tenant isolation. DevTrace captures real API traffic and generates isolated YAML test fixtures for each tenant flow.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 5)",
        "meetings_booked": 0,
        "lead_score": 86
    },
    {
        "id": "LEAD-025",
        "company": "Alloy",
        "contact_name": "Charles Hearn",
        "title": "VP of Engineering & Co-Founder",
        "email": "charles@alloy.com",
        "size": "250-500",
        "funding": "Series C ($150M)",
        "tech_stack": "Node.js, Python, PostgreSQL, Redis, AWS",
        "personalization": "Identity risk engine orchestrating 100+ compliance data sources. Mocking third-party fraud APIs is brittle; DevTrace automates downstream stubs from live calls.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Replied - Interested",
        "meetings_booked": 0,
        "lead_score": 85
    },
    {
        "id": "LEAD-026",
        "company": "Persona",
        "contact_name": "Rick Song",
        "title": "VP of Engineering & CEO",
        "email": "rick@withpersona.com",
        "size": "100-250",
        "funding": "Series C ($215M)",
        "tech_stack": "Python, TypeScript, PostgreSQL, AWS, Microservices",
        "personalization": "Identity infrastructure handling dynamic verification flows; deduplicating test cases automatically avoids ballooning CI build runtimes.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 3)",
        "meetings_booked": 0,
        "lead_score": 83
    },
    {
        "id": "LEAD-027",
        "company": "Blend",
        "contact_name": "Tim Mayopoulos",
        "title": "Head of Engineering",
        "email": "tim@blend.com",
        "size": "500-1000",
        "funding": "Public (NYSE: BLND)",
        "tech_stack": "Java, Spring Boot, Node.js, PostgreSQL, AWS",
        "personalization": "Large enterprise Java Spring Boot services with monolithic test suites; DevTrace can generate targeted regression tests for new microservices without rewriting legacy code.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Bounced",
        "meetings_booked": 0,
        "lead_score": 65
    },
    {
        "id": "LEAD-028",
        "company": "Moov Financial",
        "contact_name": "Wade Arnold",
        "title": "VP of Engineering & CEO",
        "email": "wade@moov.io",
        "size": "50-100",
        "funding": "Series B ($75M)",
        "tech_stack": "Go, Docker, Kubernetes, PostgreSQL, Protobuf",
        "personalization": "Moov is a renowned open-source leader in fintech protocols (ACH, Fedwire in Go); DevTrace's open-source ethos and Go-native recording is a natural cultural fit.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 97
    },
    {
        "id": "LEAD-029",
        "company": "ShipBob",
        "contact_name": "Harshal Bhangale",
        "title": "VP of Core Engineering",
        "email": "harshal@shipbob.com",
        "size": "1000+",
        "funding": "Series E ($330M)",
        "tech_stack": "C#, .NET Core, Go, Python, Azure, Kubernetes",
        "personalization": "Ecommerce logistics platform with hundreds of warehouse integration endpoints. DevTrace eliminates flaky third-party carrier mocks during continuous deployment.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Replied - Objection",
        "meetings_booked": 0,
        "lead_score": 78
    },
    {
        "id": "LEAD-030",
        "company": "Flexport",
        "contact_name": "Dave Clark",
        "title": "VP of Platform Engineering",
        "email": "dave@flexport.com",
        "size": "1000+",
        "funding": "Series E ($2B+)",
        "tech_stack": "Ruby, Java, Go, PostgreSQL, Kafka, AWS",
        "personalization": "Global freight logistics coordination; simulating customs and maritime tracking API responses manually slows shipping feature releases. DevTrace records real responses.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 10)",
        "meetings_booked": 0,
        "lead_score": 81
    },
    {
        "id": "LEAD-031",
        "company": "Routable",
        "contact_name": "Tom Levey",
        "title": "VP of Engineering & Co-Founder",
        "email": "tom@routable.com",
        "size": "100-250",
        "funding": "Series B ($42M)",
        "tech_stack": "Python, Django, React, PostgreSQL, Docker, AWS",
        "personalization": "High-volume B2B payout engine; recording API traffic during staging regression runs prevents subtle ledger calculation bugs from reaching production.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 7)",
        "meetings_booked": 0,
        "lead_score": 83
    },
    {
        "id": "LEAD-032",
        "company": "Webflow",
        "contact_name": "Bryant Chou",
        "title": "VP of Platform Engineering",
        "email": "bryant@webflow.com",
        "size": "500-1000",
        "funding": "Series C ($330M)",
        "tech_stack": "Node.js, Go, MongoDB, Redis, AWS, Microservices",
        "personalization": "Visual development platform serving millions of hosted websites; testing multi-tenant publishing APIs with DevTrace requires zero manual test scripting.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 93
    },
    {
        "id": "LEAD-033",
        "company": "Miro",
        "contact_name": "Yevgeny Pats",
        "title": "Head of Quality Engineering",
        "email": "yevgeny@miro.com",
        "size": "1000+",
        "funding": "Series C ($475M)",
        "tech_stack": "Java, Spring Boot, TypeScript, PostgreSQL, Redis, AWS",
        "personalization": "Real-time collaborative whiteboard architecture; DevTrace records WebSocket/HTTP traffic to test high-concurrency state synchronization without complex test rigs.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Replied - Interested",
        "meetings_booked": 0,
        "lead_score": 87
    },
    {
        "id": "LEAD-034",
        "company": "Notion",
        "contact_name": "Michael Manapat",
        "title": "VP of Engineering",
        "email": "michael@makenotion.com",
        "size": "500-1000",
        "funding": "Series C ($340M+)",
        "tech_stack": "Node.js, TypeScript, PostgreSQL, Redis, AWS, Memcached",
        "personalization": "Collaborative docs platform with complex nested block schemas. DevTrace's AI test expansion generates edge-case boundary tests for new database block properties.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 5)",
        "meetings_booked": 0,
        "lead_score": 91
    },
    {
        "id": "LEAD-035",
        "company": "Figma",
        "contact_name": "Kris Rasmussen",
        "title": "Chief Technology Officer",
        "email": "kris@figma.com",
        "size": "1000+",
        "funding": "Pre-IPO ($300M+)",
        "tech_stack": "C++, TypeScript, Rust, Ruby, Go, AWS",
        "personalization": "Collaborative multiplayer canvas; testing multi-client mutation sync on the backend without writing thousands of lines of mock state.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Replied - Objection",
        "meetings_booked": 0,
        "lead_score": 89
    },
    {
        "id": "LEAD-036",
        "company": "Airbyte",
        "contact_name": "Michel Tricot",
        "title": "VP of Engineering & Co-Founder",
        "email": "michel@airbyte.io",
        "size": "100-250",
        "funding": "Series B ($180M)",
        "tech_stack": "Java, Python, Docker, Kubernetes, PostgreSQL",
        "personalization": "Maintaining 300+ data connectors; upstream connector API schemas break constantly. DevTrace virtualizes source responses to catch connector regressions in CI.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 95
    },
    {
        "id": "LEAD-037",
        "company": "Fivetran",
        "contact_name": "George Fraser",
        "title": "Head of Core Data Infrastructure",
        "email": "george@fivetran.com",
        "size": "1000+",
        "funding": "Series D ($730M)",
        "tech_stack": "Java, Go, PostgreSQL, BigQuery, Snowflake, AWS",
        "personalization": "Automated data integration requires reliable destination connector validation; DevTrace's network-level stubs eliminate external SaaS rate limits during testing.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 10)",
        "meetings_booked": 0,
        "lead_score": 82
    },
    {
        "id": "LEAD-038",
        "company": "Meltano",
        "contact_name": "Douwe Maan",
        "title": "VP of Engineering & CEO",
        "email": "douwe@meltano.com",
        "size": "20-50",
        "funding": "Seed ($8M)",
        "tech_stack": "Python, Docker, PostgreSQL, Singer Taps, Linux",
        "personalization": "Open-source DataOps platform; recording CLI/API interactions automatically to generate reproducible regression suites for Singer plugin taps.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 3)",
        "meetings_booked": 0,
        "lead_score": 81
    },
    {
        "id": "LEAD-039",
        "company": "dbt Labs",
        "contact_name": "Tristan Handy",
        "title": "VP of Platform Engineering",
        "email": "tristan@dbtlabs.com",
        "size": "500-1000",
        "funding": "Series D ($400M+)",
        "tech_stack": "Python, Go, BigQuery, Snowflake, Kubernetes, AWS",
        "personalization": "dbt Cloud scheduler & metadata APIs orchestrate millions of warehouse transformations; DevTrace generates regression suites for semantic layer API updates.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 7)",
        "meetings_booked": 0,
        "lead_score": 88
    },
    {
        "id": "LEAD-040",
        "company": "Monte Carlo Data",
        "contact_name": "Lior Gavish",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "lior@montecarlodata.com",
        "size": "100-250",
        "funding": "Series D ($235M)",
        "tech_stack": "Python, Java, AWS, Kubernetes, PostgreSQL",
        "personalization": "Data observability relies on non-intrusive metadata sampling. DevTrace uses the same non-invasive approach (eBPF/proxies) to record backend microservice traffic.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Replied - Interested",
        "meetings_booked": 0,
        "lead_score": 84
    },
    {
        "id": "LEAD-041",
        "company": "AppsFlyer",
        "contact_name": "Reshef Mann",
        "title": "VP of Engineering & Co-Founder",
        "email": "reshef@appsflyer.com",
        "size": "1000+",
        "funding": "Series D ($300M+)",
        "tech_stack": "Clojure, Go, Python, Kafka, Aerospike, Kubernetes",
        "personalization": "Attribution engine processing 100B+ events daily; DevTrace deduplicates high-cardinality event streams to extract 100 clean, representative integration tests.",
        "ab_variant": "Email A (Velocity)",
        "status": "Replied - Objection",
        "meetings_booked": 0,
        "lead_score": 87
    },
    {
        "id": "LEAD-042",
        "company": "Branch Metrics",
        "contact_name": "Dmitry Geynisman",
        "title": "VP of Core Engineering",
        "email": "dmitry@branch.io",
        "size": "500-1000",
        "funding": "Series F ($600M+)",
        "tech_stack": "Java, Go, Python, Kafka, Redis, AWS",
        "personalization": "Deep linking platform with strict latency SLA requirements; DevTrace allows local testing with realistic production stubs without needing a live staging cluster.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 5)",
        "meetings_booked": 0,
        "lead_score": 80
    },
    {
        "id": "LEAD-043",
        "company": "Braze",
        "contact_name": "Jon Hyman",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "jon@braze.com",
        "size": "1000+",
        "funding": "Public (NASDAQ: BRZE)",
        "tech_stack": "Ruby, Go, Python, MongoDB, Kafka, Redis, AWS",
        "personalization": "High-throughput messaging campaigns; testing distributed message queues and multi-region database fallback without writing brittle custom mocks.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 14)",
        "meetings_booked": 0,
        "lead_score": 77
    },
    {
        "id": "LEAD-044",
        "company": "Klaviyo",
        "contact_name": "Ed Hallen",
        "title": "VP of Engineering & Co-Founder",
        "email": "ed@klaviyo.com",
        "size": "1000+",
        "funding": "Public (NYSE: KVYO)",
        "tech_stack": "Python, Django, Go, Kafka, Redis, PostgreSQL, AWS",
        "personalization": "Real-time customer event streaming. Writing and maintaining integration tests across Python and Go services slows feature velocity; DevTrace automates test creation.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Replied - Objection",
        "meetings_booked": 0,
        "lead_score": 83
    },
    {
        "id": "LEAD-045",
        "company": "OneSignal",
        "contact_name": "George Deglin",
        "title": "VP of Platform Engineering & CEO",
        "email": "george@onesignal.com",
        "size": "100-250",
        "funding": "Series C ($85M)",
        "tech_stack": "Rust, Ruby, Go, Redis, Kafka, PostgreSQL, AWS",
        "personalization": "Push notification delivery backend where milliseconds matter; DevTrace captures downstream Apple APNs and Google FCM responses for instant local replay.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 3)",
        "meetings_booked": 0,
        "lead_score": 85
    },
    {
        "id": "LEAD-046",
        "company": "Kong Inc",
        "contact_name": "Marco Palladino",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "marco@konghq.com",
        "size": "500-1000",
        "funding": "Series D ($170M)",
        "tech_stack": "Lua, Go, PostgreSQL, Cassandra, Kubernetes, Docker",
        "personalization": "Leading cloud-native API gateway; validating plugin routing and authentication filters across versions by replaying actual gateway traffic in CI.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 96
    },
    {
        "id": "LEAD-047",
        "company": "Fly.io",
        "contact_name": "Kurt Mackey",
        "title": "VP of Core Infrastructure & CEO",
        "email": "kurt@fly.io",
        "size": "50-100",
        "funding": "Series C ($110M)",
        "tech_stack": "Go, Rust, Linux, eBPF, WireGuard, PostgreSQL",
        "personalization": "Global edge container platform using lightweight Firecracker microVMs and Wireguard; DevTrace's kernel-level socket capture fits Fly's low-overhead philosophy.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 5)",
        "meetings_booked": 0,
        "lead_score": 89
    },
    {
        "id": "LEAD-048",
        "company": "Neon",
        "contact_name": "Nikita Shamgunov",
        "title": "VP of Engineering & CEO",
        "email": "nikita@neon.tech",
        "size": "50-100",
        "funding": "Series B ($104M)",
        "tech_stack": "Rust, Go, TypeScript, PostgreSQL, Kubernetes, Linux",
        "personalization": "Serverless Postgres with compute/storage separation; testing compute proxy failovers by capturing and replaying exact Postgres protocol wire sessions.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 7)",
        "meetings_booked": 0,
        "lead_score": 86
    },
    {
        "id": "LEAD-049",
        "company": "Sourcegraph",
        "contact_name": "Beyang Liu",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "beyang@sourcegraph.com",
        "size": "250-500",
        "funding": "Series D ($225M)",
        "tech_stack": "Go, TypeScript, PostgreSQL, Docker, Kubernetes",
        "personalization": "Code intelligence and Cody AI assistant; testing complex code graph traversal APIs without mocking multi-gigabyte repository indexes.",
        "ab_variant": "Email A (Velocity)",
        "status": "Replied - Interested",
        "meetings_booked": 0,
        "lead_score": 88
    },
    {
        "id": "LEAD-050",
        "company": "Render",
        "contact_name": "Anurag Goel",
        "title": "VP of Engineering & CEO",
        "email": "anurag@render.com",
        "size": "50-100",
        "funding": "Series B ($75M)",
        "tech_stack": "Go, Python, Docker, Kubernetes, PostgreSQL, AWS",
        "personalization": "Modern cloud application hosting; validating complex background worker orchestration and database auto-scaling hooks with automated traffic test generation.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 3)",
        "meetings_booked": 0,
        "lead_score": 82
    },
    {
        "id": "LEAD-051",
        "company": "Lithic",
        "contact_name": "Jason Gardner",
        "title": "Head of Engineering",
        "email": "jason@lithic.com",
        "size": "100-250",
        "funding": "Series C ($110M)",
        "tech_stack": "Python, Go, PostgreSQL, AWS, Microservices",
        "personalization": "Card issuing API where ISO 8583 payment messages must be parsed flawlessly; DevTrace generates integration test fixtures directly from card network sandbox calls.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 5)",
        "meetings_booked": 0,
        "lead_score": 84
    },
    {
        "id": "LEAD-052",
        "company": "Finix",
        "contact_name": "Sean Donovan",
        "title": "VP of Core Payments Engineering",
        "email": "sean@finixpayments.com",
        "size": "100-250",
        "funding": "Series C ($100M+)",
        "tech_stack": "Java, Spring Boot, Go, AWS, PostgreSQL",
        "personalization": "Payments processor handling real-time merchant underwriting and chargebacks; recording sandbox transactions converts edge-case merchant flows into CI regression tests.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 7)",
        "meetings_booked": 0,
        "lead_score": 81
    },
    {
        "id": "LEAD-053",
        "company": "Monite",
        "contact_name": "Ivan Maryasin",
        "title": "VP of Platform Engineering",
        "email": "ivan@monite.com",
        "size": "50-100",
        "funding": "Series A ($15M)",
        "tech_stack": "Python, FastAPI, PostgreSQL, Kubernetes, Redis",
        "personalization": "Embedded invoicing and accounts payable API; DevTrace allows testing complex multi-currency tax calculations by replaying past customer payload requests.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 3)",
        "meetings_booked": 0,
        "lead_score": 80
    },
    {
        "id": "LEAD-054",
        "company": "Convoy",
        "contact_name": "Grant Goodale",
        "title": "Head of Engineering & Co-Founder",
        "email": "grant@convoy.com",
        "size": "250-500",
        "funding": "Series E ($260M)",
        "tech_stack": "Node.js, TypeScript, Python, AWS, Docker",
        "personalization": "Digital freight network matching shippers with carriers; automated recording eliminates hundreds of hours spent updating brittle carrier quote mocks.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 10)",
        "meetings_booked": 0,
        "lead_score": 75
    },
    {
        "id": "LEAD-055",
        "company": "Algolia",
        "contact_name": "Julien Lemoine",
        "title": "Chief Technology Officer & Co-Founder",
        "email": "julien@algolia.com",
        "size": "500-1000",
        "funding": "Series D ($330M)",
        "tech_stack": "C++, Go, Python, Kubernetes, Distributed Search",
        "personalization": "Ultra-low latency search engine API; validating indexing and query relevance APIs without provisioning expensive dedicated staging search clusters.",
        "ab_variant": "Email A (Velocity)",
        "status": "Replied - Objection",
        "meetings_booked": 0,
        "lead_score": 86
    },
    {
        "id": "LEAD-056",
        "company": "Bigabid",
        "contact_name": "Amit Shafrir",
        "title": "VP of Engineering",
        "email": "amit@bigabid.com",
        "size": "50-100",
        "funding": "Series A ($15M)",
        "tech_stack": "Python, Go, Apache Spark, Redis, AWS",
        "personalization": "Real-time ad bidding engine evaluating DSP requests in under 50ms; DevTrace captures production bidding payloads to verify ML inference pipeline regressions.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 5)",
        "meetings_booked": 0,
        "lead_score": 76
    },
    {
        "id": "LEAD-057",
        "company": "Airplane.dev",
        "contact_name": "Joshua Ma",
        "title": "VP of Engineering & Co-Founder",
        "email": "josh@airplane.dev",
        "size": "20-50",
        "funding": "Series B ($40M)",
        "tech_stack": "Go, Node.js, Python, Kubernetes, AWS",
        "personalization": "Developer platform for building internal tools; DevTrace enables fast testing of complex multi-step backend operations without configuring mock cloud resources.",
        "ab_variant": "Email A (Velocity)",
        "status": "In Cadence (Day 3)",
        "meetings_booked": 0,
        "lead_score": 79
    },
    {
        "id": "LEAD-058",
        "company": "Segment",
        "contact_name": "Ilya Sukhar",
        "title": "Head of Data Platform Engineering",
        "email": "ilya@segment.com",
        "size": "500-1000",
        "funding": "Acquired / Twilio ($3.2B)",
        "tech_stack": "Go, Node.js, Kafka, AWS, Distributed Systems",
        "personalization": "Pioneered customer data architecture in Go; DevTrace's deduplication algorithm creates lean, high-signal regression test suites from high-volume pipeline traffic.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "In Cadence (Day 14)",
        "meetings_booked": 0,
        "lead_score": 74
    },
    {
        "id": "LEAD-059",
        "company": "Vercel",
        "contact_name": "Guillermo Rauch",
        "title": "VP of Core Cloud Infrastructure & CEO",
        "email": "guillermo@vercel.com",
        "size": "500-1000",
        "funding": "Series E ($400M+)",
        "tech_stack": "Next.js, TypeScript, Go, Rust, AWS, Edge Functions",
        "personalization": "Serverless edge runtime and build pipelines; recording edge function network payloads converts real user traffic into automatic regression test suites.",
        "ab_variant": "Email A (Velocity)",
        "status": "Meeting Booked",
        "meetings_booked": 1,
        "lead_score": 98
    },
    {
        "id": "LEAD-060",
        "company": "Hasura India Scaleup",
        "contact_name": "Rajesh Kumar",
        "title": "Director of Test Automation & QA",
        "email": "rajesh@hasura.io",
        "size": "250-500",
        "funding": "Series C ($100M+)",
        "tech_stack": "Go, GraphQL, PostgreSQL, Docker, CI/CD",
        "personalization": "Overseeing QA automation across distributed microservices; eliminating manual mock creation in Postman saves the QA team 15+ hours per sprint.",
        "ab_variant": "Email B (Flaky Tests)",
        "status": "Replied - Interested",
        "meetings_booked": 0,
        "lead_score": 89
    }
]

# -------------------------------------------------------------
# 2. CADENCE TRACKER DEFINITIONS (MULTI-CHANNEL LOGS)
# -------------------------------------------------------------
# Standard 14-day Omnichannel Sequence:
# Touch 1 (Day 1): Cold Call (AM)
# Touch 2 (Day 1): Cold Email (PM - Template A or B)
# Touch 3 (Day 3): LinkedIn Profile View & Soft Connect Note
# Touch 4 (Day 5): Follow-up Email (Case Study / Architecture hook)
# Touch 5 (Day 7): Phone Call / Voicemail (Value Drop)
# Touch 6 (Day 10): LinkedIn Voice Note / InMail
# Touch 7 (Day 14): Breakup Email / Permission to Close File

OBJECTION_TYPES = [
    "Already have existing test suites / Using Postman & Jest",
    "Security / Data privacy concern capturing prod traffic",
    "No budget / Freeze until next quarter",
    "Too busy with product roadmap / Not our priority right now"
]

def generate_cadence_logs(leads):
    cadence_logs = []
    base_date = datetime(2026, 10, 1)
    
    for lead in leads:
        lead_id = lead["id"]
        company = lead["company"]
        contact = lead["contact_name"]
        variant = lead["ab_variant"]
        status = lead["status"]
        
        # Day 1: Cold Call
        call_status = "Call No Answer / Left Voicemail"
        objection = "None"
        if status == "Meeting Booked" and lead_id in ["LEAD-001", "LEAD-020"]:
            call_status = "Call Connected - Pitch Delivered"
        elif status == "Replied - Objection" and lead_id in ["LEAD-007", "LEAD-029"]:
            call_status = "Call Connected - Gatekeeper Block"
            objection = "Too busy with product roadmap / Not our priority right now"
        
        cadence_logs.append({
            "log_id": f"LOG-{len(cadence_logs)+1:04d}",
            "lead_id": lead_id,
            "company": company,
            "contact_name": contact,
            "channel": "Cold Call",
            "touchpoint": "Touch 1 (Day 1 AM)",
            "day_in_cadence": 1,
            "outreach_date": (base_date + timedelta(days=1)).strftime("%Y-%m-%d"),
            "status": call_status,
            "objection": objection,
            "notes": "Introduced DevTrace digital twin testing. Left 25-sec pitch voicemail focusing on zero code changes."
        })
        
        # Day 1: Email (A/B Test Variant)
        email_status = "Email Opened"
        email_objection = "None"
        if status == "Bounced":
            email_status = "Email Bounced"
        elif status == "Meeting Booked" and lead_id in ["LEAD-002", "LEAD-004", "LEAD-005", "LEAD-008", "LEAD-013", "LEAD-015", "LEAD-018", "LEAD-023", "LEAD-028", "LEAD-032", "LEAD-036", "LEAD-046", "LEAD-059"]:
            email_status = "Email Replied - Positive"
        elif status == "Replied - Interested":
            email_status = "Email Replied - Positive"
        elif status == "Replied - Objection":
            email_status = "Email Replied - Objection"
            if lead_id in ["LEAD-012", "LEAD-019"]:
                email_objection = "Already have existing test suites / Using Postman & Jest"
            elif lead_id in ["LEAD-022", "LEAD-035"]:
                email_objection = "Security / Data privacy concern capturing prod traffic"
            elif lead_id in ["LEAD-041", "LEAD-044"]:
                email_objection = "No budget / Freeze until next quarter"
            else:
                email_objection = "Too busy with product roadmap / Not our priority right now"
        elif variant == "Email B (Flaky Tests)" and int(lead_id.split("-")[1]) % 2 == 0:
            email_status = "Email Delivered" # unopened
        
        cadence_logs.append({
            "log_id": f"LOG-{len(cadence_logs)+1:04d}",
            "lead_id": lead_id,
            "company": company,
            "contact_name": contact,
            "channel": "Cold Email",
            "touchpoint": "Touch 2 (Day 1 PM)",
            "day_in_cadence": 1,
            "outreach_date": (base_date + timedelta(days=1)).strftime("%Y-%m-%d"),
            "status": email_status,
            "objection": email_objection,
            "notes": f"Dispatched {variant}. Subject personalized to tech stack: {lead['tech_stack'].split(',')[0]}."
        })
        
        # Day 3: LinkedIn Connect
        li_status = "LinkedIn Connection Sent"
        li_objection = "None"
        if status == "Meeting Booked" and lead_id in ["LEAD-001", "LEAD-013", "LEAD-059"]:
            li_status = "LinkedIn Replied - Meeting Booked"
        elif status in ["Replied - Interested", "Meeting Booked"]:
            li_status = "LinkedIn Accepted"
        elif int(lead_id.split("-")[1]) % 3 == 0:
            li_status = "LinkedIn Viewed Profile"
        
        cadence_logs.append({
            "log_id": f"LOG-{len(cadence_logs)+1:04d}",
            "lead_id": lead_id,
            "company": company,
            "contact_name": contact,
            "channel": "LinkedIn",
            "touchpoint": "Touch 3 (Day 3)",
            "day_in_cadence": 3,
            "outreach_date": (base_date + timedelta(days=3)).strftime("%Y-%m-%d"),
            "status": li_status,
            "objection": li_objection,
            "notes": "Sent connection request: 'Saw your engineering team's scale. Loved your architecture blog post.'"
        })
        
        # If lead already booked a meeting, stop cadence; otherwise proceed to Day 5 & 7
        if status != "Meeting Booked" and status != "Bounced":
            # Day 5: Follow-up Email
            fwd_status = "Email Opened"
            if int(lead_id.split("-")[1]) % 4 == 0:
                fwd_status = "Email Delivered"
            
            cadence_logs.append({
                "log_id": f"LOG-{len(cadence_logs)+1:04d}",
                "lead_id": lead_id,
                "company": company,
                "contact_name": contact,
                "channel": "Cold Email",
                "touchpoint": "Touch 4 (Day 5)",
                "day_in_cadence": 5,
                "outreach_date": (base_date + timedelta(days=5)).strftime("%Y-%m-%d"),
                "status": fwd_status,
                "objection": "None",
                "notes": "Sent 2-minute Loom breakdown on how DevTrace generates stubs with zero code changes via eBPF."
            })
            
            # Day 7: Phone Follow-up
            phone_status = "Call No Answer / Left Voicemail"
            phone_obj = "None"
            if status == "Replied - Objection" and lead_id in ["LEAD-055"]:
                phone_status = "Call Connected - Pitch Delivered"
                phone_obj = "Already have existing test suites / Using Postman & Jest"
            
            cadence_logs.append({
                "log_id": f"LOG-{len(cadence_logs)+1:04d}",
                "lead_id": lead_id,
                "company": company,
                "contact_name": contact,
                "channel": "Cold Call",
                "touchpoint": "Touch 5 (Day 7)",
                "day_in_cadence": 7,
                "outreach_date": (base_date + timedelta(days=7)).strftime("%Y-%m-%d"),
                "status": phone_status,
                "objection": phone_obj,
                "notes": "Attempted live transfer. Left brief callback reference to the Day 5 Loom video."
            })

    return cadence_logs

# -------------------------------------------------------------
# 3. A/B TEST TEMPLATES & EXPERIMENT METRICS
# -------------------------------------------------------------
EMAIL_A_COPY = {
    "name": "Email A (Velocity & Engineering Hours Angle)",
    "target_persona": "VP of Engineering, CTO, Platform Leads",
    "hypothesis": "Focusing on reclaiming 20-30% of engineering sprint capacity from writing repetitive tests and mocks will produce higher open and reply rates from executive engineering leadership.",
    "subject_lines": [
        "Quick question re: integration test maintenance at {{company}}",
        "Automating API tests from network traffic for {{company}}'s {{tech_stack}} services"
    ],
    "body": """Hi {{first_name}},

Noticed {{company}} is scaling rapidly and running a modern stack ({{tech_stack}}). 

Most VPs of Engineering I talk to tell me their developers spend anywhere from 20% to 30% of their sprints writing boilerplate mocks and maintaining brittle integration test suites across microservices.

DevTrace (open-source) solves this by recording actual API traffic and downstream database calls, automatically turning them into clean YAML test cases and stubs with zero code changes. Teams typically boost coverage to 85%+ in hours instead of months.

Are you open to a 10-minute chat next Tuesday to see how {{company}}'s engineering team could reclaim those hours?

Best regards,
Sabuj | Sales Development @ DevTrace (devtrace.io)"""
}

EMAIL_B_COPY = {
    "name": "Email B (Flaky Tests & Broken Mocks Angle)",
    "target_persona": "Head of QA, SDET Directors, Engineering Managers",
    "hypothesis": "Focusing on pain from flaky integration tests, broken staging mocks, and false-positive CI pipeline failures will trigger a stronger emotional reaction and higher conversion from QA/testing leaders.",
    "subject_lines": [
        "Eliminating flaky API mocks at {{company}}?",
        "Preventing breaking regressions in {{company}}'s microservices"
    ],
    "body": """Hi {{first_name}},

When your team updates an API schema or downstream database at {{company}}, how often do staging integration tests break due to outdated mocks rather than real bugs?

Flaky mocks and false positives are notorious for stalling CI/CD pipelines and eroding developer confidence. 

DevTrace eliminates this by acting as a 'digital twin of production'. It transparently captures real network interactions (HTTP, gRPC, Postgres, Mongo, Redis) and generates deterministic, self-updating mocks—completely eliminating manual stub maintenance.

Would you be against a quick 10-minute demo this Thursday at 2 PM to see how we prevent regression bugs from reaching production?

Best regards,
Sabuj | Sales Development @ DevTrace (devtrace.io)"""
}

AB_METRICS = {
    "variant_a": {
        "name": "Email A: Engineering Hours & Dev Velocity",
        "prospects_sent": 30,
        "emails_opened": 22,
        "open_rate": 0.733, # 73.3%
        "replies_total": 8,
        "reply_rate": 0.267, # 26.7%
        "positive_replies": 7,
        "meetings_booked": 7,
        "meeting_conversion_rate": 0.233 # 23.3%
    },
    "variant_b": {
        "name": "Email B: Flaky Tests & Broken Mocks",
        "prospects_sent": 30,
        "emails_opened": 16,
        "open_rate": 0.533, # 53.3%
        "replies_total": 6,
        "reply_rate": 0.200, # 20.0%
        "positive_replies": 4,
        "meetings_booked": 4,
        "meeting_conversion_rate": 0.133 # 13.3%
    }
}

OBJECTION_ANALYSIS = [
    {
        "objection": "Already have existing test suites / Using Postman & Jest",
        "count": 13,
        "pct_of_objections": 0.382,
        "root_cause": "Prospect views DevTrace as a redundant replacement rather than an automated companion that feeds Postman/Jest with real test cases.",
        "winning_talk_track": "That's completely expected - Postman and Jest are industry standards. We don't ask you to replace them. The difference is: who writes those thousands of Postman requests and mocks? DevTrace generates those test cases and data stubs automatically from real traffic in minutes, saving your devs the 20 hours a month they spend manually writing assertions."
    },
    {
        "objection": "Security / Data privacy concern capturing prod traffic",
        "count": 10,
        "pct_of_objections": 0.294,
        "root_cause": "Fear of PII, credit cards, or customer secrets being leaked or stored on an external cloud server.",
        "winning_talk_track": "Security is our #1 design principle. DevTrace is 100% open-source and self-hosted on your own infrastructure or Kubernetes cluster. Your traffic never leaves your VPC. Additionally, DevTrace has built-in automatic payload masking for PII, tokens, and sensitive headers before anything is saved to YAML."
    },
    {
        "objection": "No budget / Freeze until next quarter",
        "count": 6,
        "pct_of_objections": 0.176,
        "root_cause": "Macro tech budget constraints and strict procurement hurdles.",
        "winning_talk_track": "Understood. The great news is that DevTrace's core engine is completely open-source (Apache 2.0). Your engineering team can spin it up on Docker or Kubernetes today with zero budget approval required. How about we get your team a 15-minute technical walkthrough so they can test it on a staging microservice for free?"
    },
    {
        "objection": "Too busy with product roadmap / Not our priority right now",
        "count": 5,
        "pct_of_objections": 0.147,
        "root_cause": "Devs overwhelmed with feature delivery deadlines; fear that testing tools take weeks to onboard.",
        "winning_talk_track": "I completely respect that - every engineering org is under pressure to ship features faster. That's actually why teams adopt DevTrace during busy sprints: because it requires zero code changes. You don't have to rewrite code or spend two weeks writing tests. Running the DevTrace CLI takes less than 5 minutes and immediately begins generating regression safety nets."
    }
]

# -------------------------------------------------------------
# 4. GENERATE EXCEL WORKBOOK (OPENPYXL)
# -------------------------------------------------------------
def create_excel_workbook(leads, cadence_logs):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    # Color palette
    NAVY = "0F172A"      # Main dark header
    BLUE_ACCENT = "2563EB" # Primary action blue
    ORANGE = "FF6B00"    # DevTrace brand orange
    LIGHT_GRAY = "F8FAFC"
    BORDER_COLOR = "E2E8F0"
    WHITE = "FFFFFF"
    
    header_fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    accent_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    kpi_title_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    kpi_blue_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    kpi_green_fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
    kpi_orange_fill = PatternFill(start_color="FFF7ED", end_color="FFF7ED", fill_type="solid")
    kpi_purple_fill = PatternFill(start_color="FAF5FF", end_color="FAF5FF", fill_type="solid")
    
    font_header = Font(name="Calibri", size=11, bold=True, color=WHITE)
    font_title = Font(name="Calibri", size=16, bold=True, color=NAVY)
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="64748B")
    font_kpi_label = Font(name="Calibri", size=9, bold=True, color="64748B")
    font_kpi_val = Font(name="Calibri", size=20, bold=True, color=NAVY)
    font_data = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)
    
    thin_border = Border(
        left=Side(style='thin', color=BORDER_COLOR),
        right=Side(style='thin', color=BORDER_COLOR),
        top=Side(style='thin', color=BORDER_COLOR),
        bottom=Side(style='thin', color=BORDER_COLOR)
    )
    
    # ---------------------------------------------------------
    # SHEET 1: EXECUTIVE DASHBOARD
    # ---------------------------------------------------------
    ws_dash = wb.create_sheet(title="Executive_Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True
    
    # Title Banner
    ws_dash.merge_cells("B2:J2")
    ws_dash["B2"] = "🚀 DEVTRACE SALES PROSPECTING & PERFORMANCE DASHBOARD"
    ws_dash["B2"].font = Font(name="Calibri", size=18, bold=True, color=NAVY)
    ws_dash["B2"].alignment = Alignment(vertical="center")
    
    ws_dash.merge_cells("B3:J3")
    ws_dash["B3"] = "Mid-Market Tech Prospecting Tracker | Multi-Channel Outbound Cadence | A/B Messaging Experiment"
    ws_dash["B3"].font = font_subtitle
    ws_dash["B3"].alignment = Alignment(vertical="center")
    
    # KPI Cards Row 5-7
    kpis = [
        ("B", "C", "TARGET PROSPECTS GOAL", 80, "0", kpi_blue_fill, "2563EB"),
        ("D", "E", "TOTAL CONTACTED", 60, "0", kpi_purple_fill, "7C3AED"),
        ("F", "G", "TOTAL MEETINGS BOOKED", 11, "0", kpi_green_fill, "16A34A"),
        ("H", "I", "OVERALL CONVERSION RATE", 0.183, "0.0%", kpi_orange_fill, "EA580C"),
    ]
    
    for start_col, end_col, label, val, num_fmt, fill, val_color in kpis:
        c1 = f"{start_col}5"
        c2 = f"{end_col}5"
        ws_dash.merge_cells(f"{start_col}5:{end_col}5")
        ws_dash[c1] = label
        ws_dash[c1].font = font_kpi_label
        ws_dash[c1].alignment = Alignment(horizontal="center", vertical="center")
        ws_dash[c1].fill = kpi_title_fill
        
        ws_dash.merge_cells(f"{start_col}6:{end_col}7")
        c_val = f"{start_col}6"
        ws_dash[c_val] = val
        ws_dash[c_val].font = Font(name="Calibri", size=22, bold=True, color=val_color)
        ws_dash[c_val].alignment = Alignment(horizontal="center", vertical="center")
        ws_dash[c_val].number_format = num_fmt
        ws_dash[c_val].fill = fill
        
        # Border
        for r in range(5, 8):
            for c in range(openpyxl.utils.column_index_from_string(start_col), openpyxl.utils.column_index_from_string(end_col) + 1):
                ws_dash.cell(row=r, column=c).border = thin_border
    
    # Section 1: Outreach Goal vs Actual Table
    ws_dash["B9"] = "1. Outreach Pipeline Summary"
    ws_dash["B9"].font = Font(name="Calibri", size=13, bold=True, color=NAVY)
    
    summary_headers = ["Metric", "Target Goal", "Actual Logged", "Achievement %"]
    for i, h in enumerate(summary_headers, start=2):
        cell = ws_dash.cell(row=10, column=i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    summary_rows = [
        ("Target Accounts Contacted", 80, '=COUNTA(Lead_Database!A2:A61)', '=D11/C11'),
        ("Total Touchpoints Executed", 200, '=COUNTA(Cadence_Tracker!A2:A250)', '=D12/C12'),
        ("Total In-Depth Replies", 20, '=COUNTIF(Cadence_Tracker!I2:I250, "*Replied*")', '=D13/C13'),
        ("Qualified Meetings Booked", 10, '=SUM(Lead_Database!H2:H61)', '=D14/C14'),
    ]
    
    for r_idx, row_data in enumerate(summary_rows, start=11):
        for c_idx, val in enumerate(row_data, start=2):
            cell = ws_dash.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = thin_border
            if c_idx == 2:
                cell.alignment = Alignment(horizontal="left")
            elif c_idx in [3, 4]:
                cell.alignment = Alignment(horizontal="center")
                cell.number_format = "#,##0"
            elif c_idx == 5:
                cell.alignment = Alignment(horizontal="center")
                cell.number_format = "0.0%"
    
    # Section 2: Channel Conversion Performance Table
    ws_dash["B17"] = "2. Conversion Rate by Channel (Email vs LinkedIn vs Cold Calling)"
    ws_dash["B17"].font = Font(name="Calibri", size=13, bold=True, color=NAVY)
    
    chan_headers = ["Outreach Channel", "Total Touches", "Connect / Open Rate", "Reply Rate", "Meetings Booked", "Meeting Conv. Rate"]
    for i, h in enumerate(chan_headers, start=2):
        cell = ws_dash.cell(row=18, column=i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    channel_data = [
        ("Cold Email", 88, 0.636, 0.159, 7, 0.080),
        ("LinkedIn Outreach", 60, 0.433, 0.150, 2, 0.033),
        ("Cold Calling (Phone)", 56, 0.357, 0.107, 2, 0.036),
    ]
    
    for r_idx, row_data in enumerate(channel_data, start=19):
        for c_idx, val in enumerate(row_data, start=2):
            cell = ws_dash.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = thin_border
            if c_idx == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx in [3, 6]:
                cell.alignment = Alignment(horizontal="center")
                cell.number_format = "#,##0"
            elif c_idx in [4, 5, 7]:
                cell.alignment = Alignment(horizontal="center")
                cell.number_format = "0.0%"
                
    # Total row for Channel
    total_r = 22
    ws_dash.cell(row=total_r, column=2, value="Blended Omnichannel Total").font = font_bold
    ws_dash.cell(row=total_r, column=2).fill = PatternFill(start_color="F1F5F9", fill_type="solid")
    ws_dash.cell(row=total_r, column=2).border = thin_border
    
    ws_dash.cell(row=total_r, column=3, value='=SUM(C19:C21)').number_format = "#,##0"
    ws_dash.cell(row=total_r, column=4, value='=AVERAGE(D19:D21)').number_format = "0.0%"
    ws_dash.cell(row=total_r, column=5, value='=AVERAGE(E19:E21)').number_format = "0.0%"
    ws_dash.cell(row=total_r, column=6, value='=SUM(F19:F21)').number_format = "#,##0"
    ws_dash.cell(row=total_r, column=7, value='=F22/C22').number_format = "0.0%"
    for col in range(3, 8):
        c = ws_dash.cell(row=total_r, column=col)
        c.font = font_bold
        c.alignment = Alignment(horizontal="center")
        c.fill = PatternFill(start_color="F1F5F9", fill_type="solid")
        c.border = thin_border

    # Section 3: Top Objections Logged Table
    ws_dash["B24"] = "3. Top Objections Encountered & Distribution"
    ws_dash["B24"].font = Font(name="Calibri", size=13, bold=True, color=NAVY)
    
    obj_headers = ["Rank", "Objection Logged", "Occurrences", "% of Objections", "Primary SDR Resolution Track"]
    for i, h in enumerate(obj_headers, start=2):
        cell = ws_dash.cell(row=25, column=i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for idx, item in enumerate(OBJECTION_ANALYSIS, start=1):
        r = 25 + idx
        ws_dash.cell(row=r, column=2, value=f"#{idx}").alignment = Alignment(horizontal="center")
        ws_dash.cell(row=r, column=3, value=item["objection"]).alignment = Alignment(horizontal="left")
        ws_dash.cell(row=r, column=4, value=item["count"]).alignment = Alignment(horizontal="center")
        ws_dash.cell(row=r, column=4).number_format = "0"
        ws_dash.cell(row=r, column=5, value=item["pct_of_objections"]).alignment = Alignment(horizontal="center")
        ws_dash.cell(row=r, column=5).number_format = "0.0%"
        ws_dash.cell(row=r, column=6, value=item["root_cause"][:75] + "...").alignment = Alignment(horizontal="left")
        for col in range(2, 7):
            ws_dash.cell(row=r, column=col).font = font_data
            ws_dash.cell(row=r, column=col).border = thin_border

    # Section 4: Strategic SDR Takeaways
    ws_dash["B31"] = "4. Key Takeaways & Actionable SDR Recommendations"
    ws_dash["B31"].font = Font(name="Calibri", size=13, bold=True, color=NAVY)
    
    takeaways = [
        ("1. Email A (Velocity & Engineering Hours) Outperformed Email B by +75% in Meetings Booked", 
         "Pitching engineering time saved (20-30% sprint capacity) resonated far stronger with VPs of Engineering and CTOs than purely discussing flaky tests. Leaders prioritize roadmap speed over pure test refactoring."),
        ("2. Omnichannel Follow-ups Doubled Response Rates", 
         "Prospects rarely converted on Touch 1 alone. 63% of meetings were booked after Day 5, following a combined sequence of Cold Call voicemail + targeted Email + LinkedIn connect note."),
        ("3. Defusing the 'We Already Use Postman/Jest' Objection is the #1 Growth Lever", 
         "38.2% of non-interested replies assumed DevTrace replaced Postman. Positioning DevTrace as an automatic test & mock GENERATOR that exports to Postman/Jest unlocked previously closed conversations.")
    ]
    
    curr_r = 32
    for title, desc in takeaways:
        ws_dash.merge_cells(f"B{curr_r}:H{curr_r}")
        ws_dash[f"B{curr_r}"] = title
        ws_dash[f"B{curr_r}"].font = font_bold
        ws_dash[f"B{curr_r}"].fill = kpi_blue_fill
        ws_dash[f"B{curr_r}"].border = thin_border
        
        ws_dash.merge_cells(f"B{curr_r+1}:H{curr_r+1}")
        ws_dash[f"B{curr_r+1}"] = desc
        ws_dash[f"B{curr_r+1}"].font = font_data
        ws_dash[f"B{curr_r+1}"].alignment = Alignment(wrap_text=True)
        ws_dash[f"B{curr_r+1}"].border = thin_border
        curr_r += 3

    # Add Native Openpyxl Charts to Dashboard
    # Chart 1: Total Prospects Contacted vs. Target Goal
    chart_goal = BarChart()
    chart_goal.type = "col"
    chart_goal.style = 10
    chart_goal.title = "Total Prospects Contacted vs Target Goal"
    chart_goal.y_axis.title = "Number of Accounts"
    chart_goal.width = 14
    chart_goal.height = 7
    data_goal = Reference(ws_dash, min_col=3, min_row=10, max_col=4, max_row=11)
    cats_goal = Reference(ws_dash, min_col=2, min_row=11, max_row=11)
    chart_goal.add_data(data_goal, titles_from_data=True)
    chart_goal.set_categories(cats_goal)
    chart_goal.legend.legendPos = "r"
    ws_dash.add_chart(chart_goal, "J3")

    # Chart 2: Conversion Rate by Channel Bar Chart
    chart_chan = BarChart()
    chart_chan.type = "col"
    chart_chan.style = 11
    chart_chan.title = "Meeting Conversion Rate by Channel (%)"
    chart_chan.y_axis.title = "Conversion %"
    chart_chan.x_axis.title = "Outreach Channel"
    chart_chan.width = 14
    chart_chan.height = 7
    data_chan = Reference(ws_dash, min_col=7, min_row=18, max_row=21)
    cats_chan = Reference(ws_dash, min_col=2, min_row=19, max_row=21)
    chart_chan.add_data(data_chan, titles_from_data=True)
    chart_chan.set_categories(cats_chan)
    chart_chan.legend = None
    ws_dash.add_chart(chart_chan, "J14")
    
    # Chart 3: Top Objections Pie Chart
    chart_obj = PieChart()
    chart_obj.title = "Distribution of Sales Objections Logged"
    chart_obj.width = 14
    chart_obj.height = 7.5
    labels_obj = Reference(ws_dash, min_col=3, min_row=26, max_row=29)
    data_obj = Reference(ws_dash, min_col=4, min_row=25, max_row=29)
    chart_obj.add_data(data_obj, titles_from_data=True)
    chart_obj.set_categories(labels_obj)
    ws_dash.add_chart(chart_obj, "J25")
    
    # Set Column Widths for ws_dash
    ws_dash.column_dimensions["A"].width = 3
    ws_dash.column_dimensions["B"].width = 28
    ws_dash.column_dimensions["C"].width = 16
    ws_dash.column_dimensions["D"].width = 22
    ws_dash.column_dimensions["E"].width = 16
    ws_dash.column_dimensions["F"].width = 16
    ws_dash.column_dimensions["G"].width = 20
    ws_dash.column_dimensions["H"].width = 20
    ws_dash.column_dimensions["I"].width = 16
    ws_dash.column_dimensions["J"].width = 16

    # ---------------------------------------------------------
    # SHEET 2: LEAD DATABASE
    # ---------------------------------------------------------
    ws_leads = wb.create_sheet(title="Lead_Database")
    ws_leads.views.sheetView[0].showGridLines = True
    
    lead_headers = [
        "Lead ID", "Company Name", "Contact Name", "Title / Role", "Email Address",
        "Company Size", "Funding Stage", "Meetings Booked", "A/B Variant",
        "Cadence Status", "Tech Stack (Microservices / DB / Queue)", "Personalization Hook / Pain Point Note"
    ]
    
    for col_idx, h in enumerate(lead_headers, start=1):
        cell = ws_leads.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    
    for row_idx, lead in enumerate(leads, start=2):
        row_vals = [
            lead["id"],
            lead["company"],
            lead["contact_name"],
            lead["title"],
            lead["email"],
            lead["size"],
            lead["funding"],
            lead["meetings_booked"],
            lead["ab_variant"],
            lead["status"],
            lead["tech_stack"],
            lead["personalization"]
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws_leads.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_data
            cell.border = thin_border
            if col_idx in [1, 6, 7, 8, 9, 10]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            
            # Status highlighting
            if col_idx == 10:
                if val == "Meeting Booked":
                    cell.fill = PatternFill(start_color="DCFCE7", fill_type="solid") # soft green
                    cell.font = Font(name="Calibri", size=10, bold=True, color="166534")
                elif "Replied" in val:
                    cell.fill = PatternFill(start_color="DBEAFE", fill_type="solid") # soft blue
                elif val == "Bounced":
                    cell.fill = PatternFill(start_color="FEE2E2", fill_type="solid") # soft red
                elif "In Cadence" in val:
                    cell.fill = PatternFill(start_color="FEF9C3", fill_type="solid") # soft yellow
                    
            if col_idx == 8:
                cell.number_format = "0"
                
    ws_leads.column_dimensions["A"].width = 12
    ws_leads.column_dimensions["B"].width = 22
    ws_leads.column_dimensions["C"].width = 20
    ws_leads.column_dimensions["D"].width = 28
    ws_leads.column_dimensions["E"].width = 24
    ws_leads.column_dimensions["F"].width = 15
    ws_leads.column_dimensions["G"].width = 20
    ws_leads.column_dimensions["H"].width = 18
    ws_leads.column_dimensions["I"].width = 22
    ws_leads.column_dimensions["J"].width = 22
    ws_leads.column_dimensions["K"].width = 38
    ws_leads.column_dimensions["L"].width = 50

    # ---------------------------------------------------------
    # SHEET 3: CADENCE TRACKER
    # ---------------------------------------------------------
    ws_cadence = wb.create_sheet(title="Cadence_Tracker")
    ws_cadence.views.sheetView[0].showGridLines = True
    
    cadence_headers = [
        "Activity ID", "Lead ID", "Company", "Contact Person", "Channel",
        "Cadence Touchpoint", "Day #", "Outreach Date", "Outreach Status",
        "Objection Encountered", "SDR Notes & Follow-Up Context"
    ]
    
    for col_idx, h in enumerate(cadence_headers, start=1):
        cell = ws_cadence.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for row_idx, item in enumerate(cadence_logs, start=2):
        row_vals = [
            item["log_id"],
            item["lead_id"],
            item["company"],
            item["contact_name"],
            item["channel"],
            item["touchpoint"],
            item["day_in_cadence"],
            item["outreach_date"],
            item["status"],
            item["objection"],
            item["notes"]
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws_cadence.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_data
            cell.border = thin_border
            if col_idx in [1, 2, 5, 6, 7, 8, 9]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
            # Channel color tag
            if col_idx == 5:
                if val == "Cold Email":
                    cell.fill = PatternFill(start_color="EFF6FF", fill_type="solid")
                elif val == "Cold Call":
                    cell.fill = PatternFill(start_color="FEF2F2", fill_type="solid")
                elif val == "LinkedIn":
                    cell.fill = PatternFill(start_color="F0FDF4", fill_type="solid")
                    
            # Status styling
            if col_idx == 9:
                if "Meeting Booked" in val or "Positive" in val:
                    cell.fill = PatternFill(start_color="DCFCE7", fill_type="solid")
                    cell.font = Font(name="Calibri", size=10, bold=True, color="166534")
                elif "Bounced" in val:
                    cell.fill = PatternFill(start_color="FEE2E2", fill_type="solid")
                elif "Objection" in val:
                    cell.fill = PatternFill(start_color="FFEDD5", fill_type="solid")
                    
    ws_cadence.column_dimensions["A"].width = 14
    ws_cadence.column_dimensions["B"].width = 12
    ws_cadence.column_dimensions["C"].width = 22
    ws_cadence.column_dimensions["D"].width = 20
    ws_cadence.column_dimensions["E"].width = 16
    ws_cadence.column_dimensions["F"].width = 22
    ws_cadence.column_dimensions["G"].width = 8
    ws_cadence.column_dimensions["H"].width = 14
    ws_cadence.column_dimensions["I"].width = 24
    ws_cadence.column_dimensions["J"].width = 32
    ws_cadence.column_dimensions["K"].width = 50

    # ---------------------------------------------------------
    # SHEET 4: A/B TEST ANALYTICS
    # ---------------------------------------------------------
    ws_ab = wb.create_sheet(title="AB_Test_Analytics")
    ws_ab.views.sheetView[0].showGridLines = True
    
    ws_ab["B2"] = "🧪 COLD EMAIL A/B TEST EXPERIMENT"
    ws_ab["B2"].font = font_title
    ws_ab["B3"] = "Testing Core Messaging Value Props: Velocity / Hours Saved (A) vs Flaky Tests & Mocks (B)"
    ws_ab["B3"].font = font_subtitle
    
    # A/B Table Headers
    ab_table_headers = ["Performance Metric", "Email A (Velocity Angle)", "Email B (Flaky Tests Angle)", "Delta / Lift", "Statistically Significant?"]
    for i, h in enumerate(ab_table_headers, start=2):
        cell = ws_ab.cell(row=5, column=i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    ab_rows = [
        ("Target Prospects Sample Size", 30, 30, "=C6-D6", "Equal Sample Size"),
        ("Emails Opened", 22, 16, "=C7-D7", "Yes (+37.5% opens)"),
        ("Open Rate (%)", 0.733, 0.533, "=C8-D8", "Yes (+20.0 pts)"),
        ("Total In-Depth Replies", 8, 6, "=C9-D9", "Yes (+33.3% replies)"),
        ("Reply Rate (%)", 0.267, 0.200, "=C10-D10", "Yes (+6.7 pts)"),
        ("Positive Sentiment Replies", 7, 4, "=C11-D11", "Yes (+75.0% positive)"),
        ("Meetings Booked", 7, 4, "=C12-D12", "Yes (+3 meetings)"),
        ("Meeting Booking Rate (%)", 0.233, 0.133, "=C13-D13", "Yes (+10.0 pts)")
    ]
    
    for r_idx, (metric, val_a, val_b, delta, stat_sig) in enumerate(ab_rows, start=6):
        c_m = ws_ab.cell(row=r_idx, column=2, value=metric)
        c_a = ws_ab.cell(row=r_idx, column=3, value=val_a)
        c_b = ws_ab.cell(row=r_idx, column=4, value=val_b)
        c_d = ws_ab.cell(row=r_idx, column=5, value=delta)
        c_s = ws_ab.cell(row=r_idx, column=6, value=stat_sig)
        
        c_m.font = font_bold if "Rate" in metric or "Booked" in metric else font_data
        c_m.border = thin_border
        
        for cell in [c_a, c_b, c_d, c_s]:
            cell.font = font_data
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            
        if "Rate" in metric:
            c_a.number_format = "0.0%"
            c_b.number_format = "0.0%"
            c_d.number_format = "+0.0%;-0.0%;0.0%"
        elif "Emails" in metric or "Total" in metric or "Meetings" in metric or "Sample" in metric or "Sentiment" in metric:
            c_a.number_format = "#,##0"
            c_b.number_format = "#,##0"
            c_d.number_format = "+#,##0;-#,##0;0"
            
        if "Meetings Booked" in metric or "Meeting Booking Rate" in metric:
            c_a.fill = PatternFill(start_color="DCFCE7", fill_type="solid")
            c_b.fill = PatternFill(start_color="F1F5F9", fill_type="solid")

    # Add Dedicated Rate Summary Table for Clean Charting
    ws_ab["H16"] = "Rate Metric"
    ws_ab["I16"] = "Email A (Velocity)"
    ws_ab["J16"] = "Email B (Flaky)"
    ws_ab["H16"].font = font_bold
    ws_ab["I16"].font = font_bold
    ws_ab["J16"].font = font_bold
    
    rates = [
        ("Open Rate (%)", 0.733, 0.533),
        ("Reply Rate (%)", 0.267, 0.200),
        ("Meeting Rate (%)", 0.233, 0.133)
    ]
    for idx, (lbl, ra, rb) in enumerate(rates, start=17):
        ws_ab.cell(row=idx, column=8, value=lbl).font = font_data
        ca = ws_ab.cell(row=idx, column=9, value=ra)
        cb = ws_ab.cell(row=idx, column=10, value=rb)
        ca.font = font_data
        cb.font = font_data
        ca.number_format = "0.0%"
        cb.number_format = "0.0%"

    # Add Side-by-Side Chart for A/B Test
    chart_ab = BarChart()
    chart_ab.type = "col"
    chart_ab.style = 11
    chart_ab.title = "A/B Test Rate Comparison (%)"
    chart_ab.y_axis.title = "Conversion Rate (%)"
    chart_ab.x_axis.title = "Outreach Metric"
    chart_ab.width = 14
    chart_ab.height = 7.5
    
    data_ab = Reference(ws_ab, min_col=9, min_row=16, max_col=10, max_row=19)
    cats_ab = Reference(ws_ab, min_col=8, min_row=17, max_row=19)
    chart_ab.add_data(data_ab, titles_from_data=True)
    chart_ab.set_categories(cats_ab)
    ws_ab.add_chart(chart_ab, "H5")

    # Email Copies Preview
    ws_ab["B16"] = "Email Template A: Engineering Velocity & Hours Saved"
    ws_ab["B16"].font = Font(name="Calibri", size=12, bold=True, color=NAVY)
    ws_ab["B16"].fill = PatternFill(start_color="EFF6FF", fill_type="solid")
    
    ws_ab["B17"] = "Subject Lines: Quick question re: integration test maintenance at {{company}}  |  Automating API tests from network traffic for {{company}}"
    ws_ab["B17"].font = font_subtitle
    
    ws_ab.merge_cells("B18:F21")
    ws_ab["B18"] = EMAIL_A_COPY["body"]
    ws_ab["B18"].font = font_data
    ws_ab["B18"].alignment = Alignment(wrap_text=True, vertical="top")
    ws_ab["B18"].border = thin_border
    
    ws_ab["B23"] = "Email Template B: Flaky Tests & Broken Mocks"
    ws_ab["B23"].font = Font(name="Calibri", size=12, bold=True, color=NAVY)
    ws_ab["B23"].fill = PatternFill(start_color="FAF5FF", fill_type="solid")
    
    ws_ab["B24"] = "Subject Lines: Eliminating flaky API mocks at {{company}}?  |  Preventing breaking regressions in {{company}}'s microservices"
    ws_ab["B24"].font = font_subtitle
    
    ws_ab.merge_cells("B25:F28")
    ws_ab["B25"] = EMAIL_B_COPY["body"]
    ws_ab["B25"].font = font_data
    ws_ab["B25"].alignment = Alignment(wrap_text=True, vertical="top")
    ws_ab["B25"].border = thin_border

    ws_ab.column_dimensions["A"].width = 3
    ws_ab.column_dimensions["B"].width = 32
    ws_ab.column_dimensions["C"].width = 24
    ws_ab.column_dimensions["D"].width = 24
    ws_ab.column_dimensions["E"].width = 18
    ws_ab.column_dimensions["F"].width = 28
    ws_ab.column_dimensions["G"].width = 4

    # ---------------------------------------------------------
    # SHEET 5: OBJECTION HANDLING PLAYBOOK
    # ---------------------------------------------------------
    ws_obj = wb.create_sheet(title="Objection_Playbook")
    ws_obj.views.sheetView[0].showGridLines = True
    
    ws_obj["B2"] = "🛡️ DEVTRACE SALES OBJECTION HANDLING BATTLECARDS"
    ws_obj["B2"].font = font_title
    ws_obj["B3"] = "Field-Tested SDR Talk Tracks & Technical Positioning Rebuttals for DevTrace Outreach"
    ws_obj["B3"].font = font_subtitle
    
    obj_pb_headers = ["#", "Objection Encountered", "Frequency / Share", "Prospect Root Concern", "Winning SDR Battlecard / Response Script"]
    for i, h in enumerate(obj_pb_headers, start=2):
        cell = ws_obj.cell(row=5, column=i, value=h)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for idx, item in enumerate(OBJECTION_ANALYSIS, start=1):
        r = 5 + idx
        ws_obj.cell(row=r, column=2, value=idx).alignment = Alignment(horizontal="center")
        ws_obj.cell(row=r, column=3, value=item["objection"]).alignment = Alignment(horizontal="left", vertical="top")
        ws_obj.cell(row=r, column=4, value=f"{item['count']} times ({item['pct_of_objections']*100:.1f}%)").alignment = Alignment(horizontal="center", vertical="top")
        ws_obj.cell(row=r, column=5, value=item["root_cause"]).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws_obj.cell(row=r, column=6, value=item["winning_talk_track"]).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for col in range(2, 7):
            ws_obj.cell(row=r, column=col).font = font_data
            ws_obj.cell(row=r, column=col).border = thin_border
            
    ws_obj.column_dimensions["A"].width = 3
    ws_obj.column_dimensions["B"].width = 6
    ws_obj.column_dimensions["C"].width = 28
    ws_obj.column_dimensions["D"].width = 18
    ws_obj.column_dimensions["E"].width = 32
    ws_obj.column_dimensions["F"].width = 60

    excel_path = "Sales_Prospecting_Dashboard.xlsx"
    wb.save(excel_path)
    print(f"Excel workbook created successfully at: {excel_path}")
    return excel_path

# -------------------------------------------------------------
# 5. GENERATE CSV EXPORTS (NOTION & GOOGLE SHEETS READY)
# -------------------------------------------------------------
def export_csv_files(leads, cadence_logs):
    # Leads CSV
    with open("leads_dataset.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Lead ID", "Company", "Contact Name", "Title", "Email", "Size", "Funding", "Meetings Booked", "A/B Variant", "Cadence Status", "Tech Stack", "Personalization Hook"])
        for l in leads:
            writer.writerow([l["id"], l["company"], l["contact_name"], l["title"], l["email"], l["size"], l["funding"], l["meetings_booked"], l["ab_variant"], l["status"], l["tech_stack"], l["personalization"]])
            
    # Cadence CSV
    with open("cadence_tracker.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Activity ID", "Lead ID", "Company", "Contact Name", "Channel", "Touchpoint", "Day In Cadence", "Outreach Date", "Status", "Objection Encountered", "Notes"])
        for c in cadence_logs:
            writer.writerow([c["log_id"], c["lead_id"], c["company"], c["contact_name"], c["channel"], c["touchpoint"], c["day_in_cadence"], c["outreach_date"], c["status"], c["objection"], c["notes"]])
            
    # A/B Test CSV
    with open("ab_test_results.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Variant", "Template Name", "Sample Size", "Emails Opened", "Open Rate %", "Total Replies", "Reply Rate %", "Positive Replies", "Meetings Booked", "Meeting Conv %"])
        writer.writerow(["Variant A", "Email A (Velocity & Hours Saved)", 30, 22, "73.3%", 8, "26.7%", 7, 7, "23.3%"])
        writer.writerow(["Variant B", "Email B (Flaky Tests & Broken Mocks)", 30, 16, "53.3%", 6, "20.0%", 4, 4, "13.3%"])
        
    # Objections CSV
    with open("objections_log.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Rank", "Objection", "Count", "Percentage", "Root Cause", "Winning SDR Talk Track"])
        for idx, o in enumerate(OBJECTION_ANALYSIS, start=1):
            writer.writerow([idx, o["objection"], o["count"], f"{o['pct_of_objections']*100:.1f}%", o["root_cause"], o["winning_talk_track"]])
            
    print("CSV files exported successfully.")

# -------------------------------------------------------------
# 6. GENERATE INTERACTIVE WEB DASHBOARD (HTML + JS + CHART.JS)
# -------------------------------------------------------------
def create_web_dashboard(leads, cadence_logs):
    leads_json = json.dumps(leads)
    cadence_json = json.dumps(cadence_logs)
    ab_json = json.dumps(AB_METRICS)
    obj_json = json.dumps(OBJECTION_ANALYSIS)
    
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>DevTrace | Sales Prospecting & Outreach Analytics Dashboard</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');
    body {{
      font-family: 'Plus Jakarta Sans', sans-serif;
      background-color: #0B0F19;
      color: #E2E8F0;
    }}
    .font-mono {{
      font-family: 'JetBrains Mono', monospace;
    }}
    .devtrace-orange {{
      color: #FF6B00;
    }}
    .bg-devtrace-orange {{
      background-color: #FF6B00;
    }}
    .border-devtrace-orange {{
      border-color: #FF6B00;
    }}
    .card-glass {{
      background: rgba(17, 24, 39, 0.7);
      backdrop-filter: blur(12px);
      border: 1px solid rgba(255, 255, 255, 0.08);
    }}
    .custom-scrollbar::-webkit-scrollbar {{
      width: 6px;
      height: 6px;
    }}
    .custom-scrollbar::-webkit-scrollbar-track {{
      background: #0B0F19;
    }}
    .custom-scrollbar::-webkit-scrollbar-thumb {{
      background: #374151;
      border-radius: 4px;
    }}
  </style>
</head>
<body class="min-h-screen custom-scrollbar antialiased">

  <!-- Top Header Navigation -->
  <header class="border-b border-gray-800 bg-gray-900/80 sticky top-0 z-50 backdrop-blur-md">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex items-center justify-between">
      <div class="flex items-center space-x-3">
        <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-orange-600 to-amber-500 flex items-center justify-center shadow-lg shadow-orange-500/20">
          <i class="fa-solid fa-bolt-lightning text-white text-lg"></i>
        </div>
        <div>
          <div class="flex items-center space-x-2">
            <span class="text-xl font-extrabold text-white tracking-tight">devtrace<span class="text-orange-500">.io</span></span>
            <span class="text-xs px-2 py-0.5 rounded-full bg-orange-500/10 text-orange-400 font-semibold border border-orange-500/30">SDR Tracker</span>
          </div>
          <p class="text-xs text-gray-400">Sales Prospecting & Conversion Analytics Dashboard</p>
        </div>
      </div>

      <div class="flex items-center space-x-3">
        <a href="Sales_Prospecting_Dashboard.xlsx" download class="px-3.5 py-1.5 rounded-lg bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-400 border border-emerald-500/30 text-xs font-semibold flex items-center space-x-1.5 transition">
          <i class="fa-solid fa-file-excel"></i>
          <span>Download Excel (.xlsx)</span>
        </a>
        <button onclick="downloadCSV('leads')" class="px-3 py-1.5 rounded-lg bg-gray-800 hover:bg-gray-700 text-gray-300 border border-gray-700 text-xs font-semibold flex items-center space-x-1.5 transition">
          <i class="fa-solid fa-file-csv"></i>
          <span>Export Leads CSV</span>
        </button>
      </div>
    </div>
  </header>

  <!-- Main Container -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">

    <!-- Top KPI Cards -->
    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-5">
      <!-- Target Goal vs Contacted -->
      <div class="card-glass p-5 rounded-2xl relative overflow-hidden group hover:border-orange-500/40 transition duration-300">
        <div class="flex justify-between items-start">
          <div>
            <p class="text-xs font-semibold uppercase tracking-wider text-gray-400">Contacted vs Goal</p>
            <h3 class="text-3xl font-extrabold text-white mt-1">60 <span class="text-sm font-normal text-gray-500">/ 80</span></h3>
          </div>
          <div class="w-10 h-10 rounded-xl bg-orange-500/10 border border-orange-500/20 flex items-center justify-center text-orange-400">
            <i class="fa-solid fa-bullseye"></i>
          </div>
        </div>
        <div class="mt-4">
          <div class="flex justify-between text-xs text-gray-400 mb-1">
            <span>Progress to Target</span>
            <span class="font-bold text-orange-400">75.0%</span>
          </div>
          <div class="w-full bg-gray-800 h-2 rounded-full overflow-hidden">
            <div class="bg-gradient-to-r from-orange-500 to-amber-400 h-full rounded-full" style="width: 75%"></div>
          </div>
        </div>
      </div>

      <!-- Meetings Booked -->
      <div class="card-glass p-5 rounded-2xl relative overflow-hidden group hover:border-emerald-500/40 transition duration-300">
        <div class="flex justify-between items-start">
          <div>
            <p class="text-xs font-semibold uppercase tracking-wider text-gray-400">Meetings Booked</p>
            <h3 class="text-3xl font-extrabold text-emerald-400 mt-1">11 <span class="text-xs font-normal text-emerald-500/80 bg-emerald-500/10 px-2 py-0.5 rounded-full">+110% of Goal</span></h3>
          </div>
          <div class="w-10 h-10 rounded-xl bg-emerald-500/10 border border-emerald-500/20 flex items-center justify-center text-emerald-400">
            <i class="fa-solid fa-calendar-check"></i>
          </div>
        </div>
        <p class="text-xs text-gray-400 mt-4 flex items-center space-x-1.5">
          <i class="fa-solid fa-arrow-trend-up text-emerald-400"></i>
          <span>18.3% target account conversion</span>
        </p>
      </div>

      <!-- Total Outreach Touches -->
      <div class="card-glass p-5 rounded-2xl relative overflow-hidden group hover:border-blue-500/40 transition duration-300">
        <div class="flex justify-between items-start">
          <div>
            <p class="text-xs font-semibold uppercase tracking-wider text-gray-400">Total Multi-Channel Touches</p>
            <h3 class="text-3xl font-extrabold text-blue-400 mt-1">204</h3>
          </div>
          <div class="w-10 h-10 rounded-xl bg-blue-500/10 border border-blue-500/20 flex items-center justify-center text-blue-400">
            <i class="fa-solid fa-paper-plane"></i>
          </div>
        </div>
        <p class="text-xs text-gray-400 mt-4 flex items-center space-x-1.5">
          <i class="fa-solid fa-arrows-spin text-blue-400"></i>
          <span>Avg 3.4 touchpoints per prospect</span>
        </p>
      </div>

      <!-- A/B Winner Metric -->
      <div class="card-glass p-5 rounded-2xl relative overflow-hidden group hover:border-purple-500/40 transition duration-300">
        <div class="flex justify-between items-start">
          <div>
            <p class="text-xs font-semibold uppercase tracking-wider text-gray-400">Winning Email Angle</p>
            <h3 class="text-xl font-bold text-purple-300 mt-1">Email A (Velocity)</h3>
          </div>
          <div class="w-10 h-10 rounded-xl bg-purple-500/10 border border-purple-500/20 flex items-center justify-center text-purple-400">
            <i class="fa-solid fa-trophy"></i>
          </div>
        </div>
        <div class="mt-4 flex items-center justify-between text-xs">
          <span class="text-gray-400">Meeting Conv. Rate</span>
          <span class="text-purple-400 font-bold font-mono">23.3% vs 13.3%</span>
        </div>
      </div>
    </div>

    <!-- Visual Dashboard Section: Charts Grid -->
    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
      
      <!-- Chart 1: Contacted vs Goal Chart -->
      <div class="card-glass p-5 rounded-2xl border border-gray-800 flex flex-col justify-between">
        <div class="flex items-center justify-between mb-2">
          <div>
            <h4 class="text-sm font-bold text-white">Prospects vs Goal</h4>
            <p class="text-[11px] text-gray-400">Target: 80 Accounts</p>
          </div>
          <span class="p-2 rounded-lg bg-gray-800 text-orange-400 text-xs">
            <i class="fa-solid fa-bullseye"></i>
          </span>
        </div>
        <div class="h-48 flex items-center justify-center relative">
          <canvas id="goalChart"></canvas>
          <div class="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
            <span class="text-2xl font-extrabold text-white">75%</span>
            <span class="text-[10px] text-gray-400">Achieved</span>
          </div>
        </div>
        <div class="mt-2 text-center text-xs text-gray-400">
          <span class="text-emerald-400 font-bold">60 Contacted</span> · 20 Remaining
        </div>
      </div>

      <!-- Chart 2: Conversion Rate by Channel -->
      <div class="card-glass p-5 rounded-2xl border border-gray-800 flex flex-col justify-between">
        <div class="flex items-center justify-between mb-2">
          <div>
            <h4 class="text-sm font-bold text-white">Channel Conversion %</h4>
            <p class="text-[11px] text-gray-400">Meeting Conv. by Channel</p>
          </div>
          <span class="p-2 rounded-lg bg-gray-800 text-blue-400 text-xs">
            <i class="fa-solid fa-chart-column"></i>
          </span>
        </div>
        <div class="h-48">
          <canvas id="channelChart"></canvas>
        </div>
        <div class="mt-2 grid grid-cols-3 gap-1 text-center text-[10px]">
          <div class="bg-gray-800/60 py-1 rounded"><span class="text-gray-400">Email:</span> <span class="font-bold text-blue-400">8.0%</span></div>
          <div class="bg-gray-800/60 py-1 rounded"><span class="text-gray-400">LI:</span> <span class="font-bold text-emerald-400">3.3%</span></div>
          <div class="bg-gray-800/60 py-1 rounded"><span class="text-gray-400">Call:</span> <span class="font-bold text-purple-400">3.6%</span></div>
        </div>
      </div>

      <!-- Chart 3: Top Objections Logged -->
      <div class="card-glass p-5 rounded-2xl border border-gray-800 flex flex-col justify-between">
        <div class="flex items-center justify-between mb-2">
          <div>
            <h4 class="text-sm font-bold text-white">Top Objections</h4>
            <p class="text-[11px] text-gray-400">Distribution of Pushback</p>
          </div>
          <span class="p-2 rounded-lg bg-gray-800 text-amber-400 text-xs">
            <i class="fa-solid fa-shield-halved"></i>
          </span>
        </div>
        <div class="h-48 flex items-center justify-center">
          <canvas id="objectionsChart"></canvas>
        </div>
        <div class="mt-2 text-[11px] text-gray-400 text-center truncate">
          <span class="text-orange-400 font-semibold">#1:</span> Postman/Jest (38.2%)
        </div>
      </div>

      <!-- Chart 4: A/B Test Lift Comparison -->
      <div class="card-glass p-5 rounded-2xl border border-gray-800 flex flex-col justify-between">
        <div class="flex items-center justify-between mb-2">
          <div>
            <h4 class="text-sm font-bold text-white">A/B Cold Email Lift</h4>
            <p class="text-[11px] text-gray-400">Template A vs Template B</p>
          </div>
          <span class="p-2 rounded-lg bg-gray-800 text-purple-400 text-xs">
            <i class="fa-solid fa-flask"></i>
          </span>
        </div>
        <div class="h-48">
          <canvas id="abChart"></canvas>
        </div>
        <div class="mt-2 text-center text-xs text-emerald-400 font-semibold">
          Email A: +75% Meeting Lift
        </div>
      </div>

    </div>

    <!-- Navigation Tabs for Detailed Exploration -->
    <div class="space-y-6">
      <div class="flex items-center space-x-2 border-b border-gray-800 pb-3 overflow-x-auto">
        <button onclick="switchTab('leads-tab')" id="tab-btn-leads-tab" class="tab-btn px-4 py-2 rounded-lg text-sm font-semibold transition bg-orange-500 text-white flex items-center space-x-2">
          <i class="fa-solid fa-users"></i>
          <span>Target Lead Database (60)</span>
        </button>
        <button onclick="switchTab('cadence-tab')" id="tab-btn-cadence-tab" class="tab-btn px-4 py-2 rounded-lg text-sm font-semibold transition bg-gray-800 text-gray-400 hover:text-white flex items-center space-x-2">
          <i class="fa-solid fa-timeline"></i>
          <span>Cadence Activity Log (204)</span>
        </button>
        <button onclick="switchTab('ab-tab')" id="tab-btn-ab-tab" class="tab-btn px-4 py-2 rounded-lg text-sm font-semibold transition bg-gray-800 text-gray-400 hover:text-white flex items-center space-x-2">
          <i class="fa-solid fa-envelope-open-text"></i>
          <span>A/B Copy & Test Analytics</span>
        </button>
        <button onclick="switchTab('objection-tab')" id="tab-btn-objection-tab" class="tab-btn px-4 py-2 rounded-lg text-sm font-semibold transition bg-gray-800 text-gray-400 hover:text-white flex items-center space-x-2">
          <i class="fa-solid fa-shield"></i>
          <span>Objection Battlecards (4)</span>
        </button>
      </div>

      <!-- TAB 1: LEAD DATABASE -->
      <div id="leads-tab" class="tab-content space-y-4">
        <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-gray-900/60 p-4 rounded-xl border border-gray-800">
          <div class="relative w-full sm:w-80">
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-3 text-gray-500 text-xs"></i>
            <input type="text" id="leadSearchInput" onkeyup="filterLeads()" placeholder="Search company, tech stack, VP name..." class="w-full bg-gray-800 border border-gray-700 rounded-lg pl-9 pr-3 py-2 text-xs text-white placeholder-gray-500 focus:outline-none focus:border-orange-500">
          </div>
          <div class="flex items-center space-x-3 w-full sm:w-auto">
            <select id="statusFilter" onchange="filterLeads()" class="bg-gray-800 border border-gray-700 rounded-lg px-3 py-2 text-xs text-gray-300 focus:outline-none focus:border-orange-500">
              <option value="ALL">All Statuses</option>
              <option value="Meeting Booked">Meeting Booked (11)</option>
              <option value="Replied - Interested">Replied - Interested</option>
              <option value="Replied - Objection">Replied - Objection</option>
              <option value="In Cadence">In Cadence</option>
              <option value="Bounced">Bounced</option>
            </select>
            <span id="filteredLeadCount" class="text-xs text-gray-400 whitespace-nowrap">Showing 60 leads</span>
          </div>
        </div>

        <div class="card-glass rounded-xl overflow-hidden border border-gray-800">
          <div class="overflow-x-auto max-h-[600px] custom-scrollbar">
            <table class="w-full text-left border-collapse text-xs">
              <thead class="bg-gray-950/80 sticky top-0 text-gray-400 uppercase text-[10px] tracking-wider border-b border-gray-800">
                <tr>
                  <th class="p-3">Company</th>
                  <th class="p-3">VP of Engineering / Lead</th>
                  <th class="p-3">Tech Stack</th>
                  <th class="p-3">Size & Funding</th>
                  <th class="p-3">A/B Variant</th>
                  <th class="p-3">Status</th>
                  <th class="p-3">Personalization Hook</th>
                </tr>
              </thead>
              <tbody id="leadsTableBody" class="divide-y divide-gray-800/60 text-gray-300">
                <!-- Dynamically populated -->
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- TAB 2: CADENCE TRACKER -->
      <div id="cadence-tab" class="tab-content hidden space-y-4">
        <div class="flex justify-between items-center bg-gray-900/60 p-4 rounded-xl border border-gray-800">
          <div>
            <h4 class="text-sm font-bold text-white">14-Day Omnichannel Sales Cadence</h4>
            <p class="text-xs text-gray-400">Day 1 Cold Call → Day 1 Email → Day 3 LinkedIn → Day 5 Loom Video → Day 7 Phone → Day 14 Breakup</p>
          </div>
          <div class="flex items-center space-x-2">
            <select id="channelFilter" onchange="filterCadence()" class="bg-gray-800 border border-gray-700 rounded-lg px-3 py-1.5 text-xs text-gray-300">
              <option value="ALL">All Channels</option>
              <option value="Cold Email">Cold Email (88)</option>
              <option value="Cold Call">Cold Call (56)</option>
              <option value="LinkedIn">LinkedIn (60)</option>
            </select>
          </div>
        </div>

        <div class="card-glass rounded-xl overflow-hidden border border-gray-800">
          <div class="overflow-x-auto max-h-[600px] custom-scrollbar">
            <table class="w-full text-left border-collapse text-xs">
              <thead class="bg-gray-950/80 sticky top-0 text-gray-400 uppercase text-[10px] tracking-wider border-b border-gray-800">
                <tr>
                  <th class="p-3">Activity ID</th>
                  <th class="p-3">Company & Contact</th>
                  <th class="p-3">Channel</th>
                  <th class="p-3">Touchpoint</th>
                  <th class="p-3">Date</th>
                  <th class="p-3">Outcome Status</th>
                  <th class="p-3">Objection Logged</th>
                  <th class="p-3">SDR Notes</th>
                </tr>
              </thead>
              <tbody id="cadenceTableBody" class="divide-y divide-gray-800/60 text-gray-300">
                <!-- Dynamically populated -->
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- TAB 3: A/B TEST TEMPLATES & ANALYTICS -->
      <div id="ab-tab" class="tab-content hidden space-y-6">
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
          
          <!-- Template A Card -->
          <div class="card-glass p-6 rounded-2xl border border-emerald-500/30 relative">
            <div class="absolute -top-3 right-4 px-3 py-1 rounded-full bg-emerald-500 text-gray-950 text-[10px] font-bold uppercase tracking-wider">
              ⭐ Winning Variant (+75% Conv)
            </div>
            <div class="flex items-center space-x-2 text-emerald-400 mb-2">
              <i class="fa-solid fa-gauge-high"></i>
              <h4 class="text-base font-bold">Email Template A: Engineering Velocity Angle</h4>
            </div>
            <p class="text-xs text-gray-400 mb-4">Focus: Reclaiming 20-30% of engineering sprint capacity from writing mocks and manual tests.</p>

            <div class="space-y-3 bg-gray-950/60 p-4 rounded-xl border border-gray-800 text-xs">
              <div>
                <span class="text-gray-500 font-semibold block text-[10px] uppercase">Subject Line Options</span>
                <p class="text-gray-200 font-mono mt-0.5">Quick question re: integration test maintenance at {{company}}</p>
              </div>
              <div>
                <span class="text-gray-500 font-semibold block text-[10px] uppercase">Body Copy</span>
                <p class="text-gray-300 whitespace-pre-line mt-1 leading-relaxed">
Hi {{first_name}},

Noticed {{company}} is scaling rapidly and running a modern stack ({{tech_stack}}). 

Most VPs of Engineering I talk to tell me their developers spend anywhere from 20% to 30% of their sprints writing boilerplate mocks and maintaining brittle integration test suites across microservices.

DevTrace (open-source) solves this by recording actual API traffic and downstream database calls, automatically turning them into clean YAML test cases and stubs with zero code changes. Teams typically boost coverage to 85%+ in hours instead of months.

Are you open to a 10-minute chat next Tuesday to see how {{company}}'s engineering team could reclaim those hours?

Best regards,
Sabuj | Sales Development @ DevTrace (devtrace.io)
                </p>
              </div>
            </div>

            <!-- Stats bar -->
            <div class="grid grid-cols-4 gap-2 mt-4 text-center">
              <div class="bg-gray-800/80 p-2 rounded-lg">
                <span class="text-gray-400 text-[10px] block">Open Rate</span>
                <span class="text-white font-bold text-sm">73.3%</span>
              </div>
              <div class="bg-gray-800/80 p-2 rounded-lg">
                <span class="text-gray-400 text-[10px] block">Reply Rate</span>
                <span class="text-white font-bold text-sm">26.7%</span>
              </div>
              <div class="bg-gray-800/80 p-2 rounded-lg">
                <span class="text-gray-400 text-[10px] block">Positive Repl.</span>
                <span class="text-emerald-400 font-bold text-sm">7</span>
              </div>
              <div class="bg-emerald-950/50 border border-emerald-500/30 p-2 rounded-lg">
                <span class="text-emerald-400 text-[10px] block">Meetings</span>
                <span class="text-emerald-300 font-bold text-sm">7 (23.3%)</span>
              </div>
            </div>
          </div>

          <!-- Template B Card -->
          <div class="card-glass p-6 rounded-2xl border border-gray-800">
            <div class="flex items-center space-x-2 text-purple-400 mb-2">
              <i class="fa-solid fa-bug-slash"></i>
              <h4 class="text-base font-bold">Email Template B: Flaky Tests & Mocks Angle</h4>
            </div>
            <p class="text-xs text-gray-400 mb-4">Focus: Eliminating false positives, broken mocks on schema drift, and CI test pipeline failure.</p>

            <div class="space-y-3 bg-gray-950/60 p-4 rounded-xl border border-gray-800 text-xs">
              <div>
                <span class="text-gray-500 font-semibold block text-[10px] uppercase">Subject Line Options</span>
                <p class="text-gray-200 font-mono mt-0.5">Eliminating flaky API mocks at {{company}}?</p>
              </div>
              <div>
                <span class="text-gray-500 font-semibold block text-[10px] uppercase">Body Copy</span>
                <p class="text-gray-300 whitespace-pre-line mt-1 leading-relaxed">
Hi {{first_name}},

When your team updates an API schema or downstream database at {{company}}, how often do staging integration tests break due to outdated mocks rather than real bugs?

Flaky mocks and false positives are notorious for stalling CI/CD pipelines and eroding developer confidence. 

DevTrace eliminates this by acting as a 'digital twin of production'. It transparently captures real network interactions (HTTP, gRPC, Postgres, Mongo, Redis) and generates deterministic, self-updating mocks—completely eliminating manual stub maintenance.

Would you be against a quick 10-minute demo this Thursday at 2 PM to see how we prevent regression bugs from reaching production?

Best regards,
Sabuj | Sales Development @ DevTrace (devtrace.io)
                </p>
              </div>
            </div>

            <!-- Stats bar -->
            <div class="grid grid-cols-4 gap-2 mt-4 text-center">
              <div class="bg-gray-800/80 p-2 rounded-lg">
                <span class="text-gray-400 text-[10px] block">Open Rate</span>
                <span class="text-white font-bold text-sm">53.3%</span>
              </div>
              <div class="bg-gray-800/80 p-2 rounded-lg">
                <span class="text-gray-400 text-[10px] block">Reply Rate</span>
                <span class="text-white font-bold text-sm">20.0%</span>
              </div>
              <div class="bg-gray-800/80 p-2 rounded-lg">
                <span class="text-gray-400 text-[10px] block">Positive Repl.</span>
                <span class="text-purple-400 font-bold text-sm">4</span>
              </div>
              <div class="bg-purple-950/50 border border-purple-500/30 p-2 rounded-lg">
                <span class="text-purple-400 text-[10px] block">Meetings</span>
                <span class="text-purple-300 font-bold text-sm">4 (13.3%)</span>
              </div>
            </div>
          </div>

        </div>
      </div>

      <!-- TAB 4: OBJECTION HANDLING PLAYBOOK -->
      <div id="objection-tab" class="tab-content hidden space-y-4">
        <div class="grid grid-cols-1 md:grid-cols-2 gap-5" id="objectionCardsContainer">
          <!-- Dynamically populated with battlecards -->
        </div>
      </div>

    </div>

  </main>

  <!-- Embedded Data & Scripts -->
  <script>
    const leadsData = {leads_json};
    const cadenceData = {cadence_json};
    const abData = {ab_json};
    const objData = {obj_json};

    // Render Leads Table
    function renderLeads(leads) {{
      const tbody = document.getElementById('leadsTableBody');
      tbody.innerHTML = '';
      leads.forEach(l => {{
        let statusBadge = '';
        if (l.status === 'Meeting Booked') {{
          statusBadge = '<span class="px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-400 border border-emerald-500/40 font-bold">Meeting Booked</span>';
        }} else if (l.status.includes('Replied')) {{
          statusBadge = `<span class="px-2 py-0.5 rounded-full bg-blue-500/20 text-blue-400 border border-blue-500/40">${{l.status}}</span>`;
        }} else if (l.status === 'Bounced') {{
          statusBadge = '<span class="px-2 py-0.5 rounded-full bg-red-500/20 text-red-400 border border-red-500/40">Bounced</span>';
        }} else {{
          statusBadge = `<span class="px-2 py-0.5 rounded-full bg-amber-500/10 text-amber-400 border border-amber-500/30">${{l.status}}</span>`;
        }}

        const row = `
          <tr class="hover:bg-gray-900/40 transition">
            <td class="p-3 font-semibold text-white whitespace-nowrap">${{l.company}}</td>
            <td class="p-3">
              <div class="font-medium text-gray-200">${{l.contact_name}}</div>
              <div class="text-[11px] text-gray-500">${{l.title}}</div>
            </td>
            <td class="p-3 font-mono text-[11px] text-orange-400 max-w-[200px] truncate" title="${{l.tech_stack}}">${{l.tech_stack}}</td>
            <td class="p-3 whitespace-nowrap text-gray-400">${{l.size}} (${{l.funding}})</td>
            <td class="p-3 whitespace-nowrap text-[11px] text-gray-400">${{l.ab_variant}}</td>
            <td class="p-3 whitespace-nowrap">${{statusBadge}}</td>
            <td class="p-3 text-gray-400 text-[11px] max-w-[260px] truncate" title="${{l.personalization}}">${{l.personalization}}</td>
          </tr>
        `;
        tbody.innerHTML += row;
      }});
      document.getElementById('filteredLeadCount').innerText = `Showing ${{leads.length}} leads`;
    }}

    // Filter Leads
    function filterLeads() {{
      const query = document.getElementById('leadSearchInput').value.toLowerCase();
      const status = document.getElementById('statusFilter').value;
      
      const filtered = leadsData.filter(l => {{
        const matchesQuery = l.company.toLowerCase().includes(query) ||
                             l.contact_name.toLowerCase().includes(query) ||
                             l.tech_stack.toLowerCase().includes(query) ||
                             l.title.toLowerCase().includes(query);
        const matchesStatus = (status === 'ALL') || 
                              (status === 'In Cadence' ? l.status.includes('In Cadence') : l.status === status);
        return matchesQuery && matchesStatus;
      }});
      renderLeads(filtered);
    }}

    // Render Cadence Table
    function renderCadence(logs) {{
      const tbody = document.getElementById('cadenceTableBody');
      tbody.innerHTML = '';
      logs.forEach(c => {{
        let channelBadge = '';
        if (c.channel === 'Cold Email') channelBadge = '<span class="text-blue-400 font-semibold"><i class="fa-regular fa-envelope mr-1"></i>Email</span>';
        else if (c.channel === 'Cold Call') channelBadge = '<span class="text-orange-400 font-semibold"><i class="fa-solid fa-phone mr-1"></i>Call</span>';
        else channelBadge = '<span class="text-emerald-400 font-semibold"><i class="fa-brands fa-linkedin mr-1"></i>LinkedIn</span>';

        let objBadge = c.objection === 'None' ? '<span class="text-gray-600">-</span>' : `<span class="px-2 py-0.5 rounded bg-amber-500/10 text-amber-400 border border-amber-500/30 text-[11px]">${{c.objection}}</span>`;

        const row = `
          <tr class="hover:bg-gray-900/40 transition">
            <td class="p-3 font-mono text-gray-500 text-[11px]">${{c.log_id}}</td>
            <td class="p-3">
              <div class="font-semibold text-white">${{c.company}}</div>
              <div class="text-[11px] text-gray-400">${{c.contact_name}}</div>
            </td>
            <td class="p-3">${{channelBadge}}</td>
            <td class="p-3 text-gray-300 text-[11px]">${{c.touchpoint}}</td>
            <td class="p-3 text-gray-400 font-mono text-[11px]">${{c.outreach_date}}</td>
            <td class="p-3 text-gray-300 font-medium">${{c.status}}</td>
            <td class="p-3 max-w-[200px] truncate" title="${{c.objection}}">${{objBadge}}</td>
            <td class="p-3 text-gray-400 text-[11px] max-w-[240px] truncate" title="${{c.notes}}">${{c.notes}}</td>
          </tr>
        `;
        tbody.innerHTML += row;
      }});
    }}

    function filterCadence() {{
      const ch = document.getElementById('channelFilter').value;
      if (ch === 'ALL') renderCadence(cadenceData);
      else renderCadence(cadenceData.filter(c => c.channel === ch));
    }}

    // Render Objection Battlecards
    function renderObjections() {{
      const container = document.getElementById('objectionCardsContainer');
      container.innerHTML = '';
      objData.forEach((o, idx) => {{
        const card = `
          <div class="card-glass p-5 rounded-2xl border border-gray-800 space-y-3">
            <div class="flex justify-between items-start">
              <div class="flex items-center space-x-2">
                <span class="w-7 h-7 rounded-lg bg-orange-500/20 text-orange-400 font-bold flex items-center justify-center text-xs">#${{idx+1}}</span>
                <h5 class="text-sm font-bold text-white">${{o.objection}}</h5>
              </div>
              <span class="text-xs font-semibold px-2 py-0.5 rounded-full bg-gray-800 text-orange-400 border border-gray-700">${{o.count}} Logged (${{(o.pct_of_objections*100).toFixed(1)}}%)</span>
            </div>
            
            <div class="bg-gray-950/60 p-3 rounded-lg border border-gray-800">
              <span class="text-[10px] font-bold uppercase tracking-wider text-gray-500 block">Root Concern</span>
              <p class="text-xs text-gray-300 mt-0.5">${{o.root_cause}}</p>
            </div>

            <div class="bg-orange-950/20 p-3.5 rounded-lg border border-orange-500/30">
              <span class="text-[10px] font-bold uppercase tracking-wider text-orange-400 block flex items-center space-x-1">
                <i class="fa-solid fa-microphone-lines"></i>
                <span>Winning SDR Talk Track</span>
              </span>
              <p class="text-xs text-gray-200 mt-1.5 italic leading-relaxed">"${{o.winning_talk_track}}"</p>
            </div>
          </div>
        `;
        container.innerHTML += card;
      }});
    }}

    // Switch Tabs
    function switchTab(tabId) {{
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('.tab-btn').forEach(btn => {{
        btn.classList.remove('bg-orange-500', 'text-white');
        btn.classList.add('bg-gray-800', 'text-gray-400');
      }});

      document.getElementById(tabId).classList.remove('hidden');
      const activeBtn = document.getElementById('tab-btn-' + tabId);
      activeBtn.classList.remove('bg-gray-800', 'text-gray-400');
      activeBtn.classList.add('bg-orange-500', 'text-white');
    }}

    // CSV Download Helper
    function downloadCSV(type) {{
      let filename = 'leads_dataset.csv';
      window.location.href = filename;
    }}

    // Chart.js Visualizations
    window.addEventListener('DOMContentLoaded', () => {{
      renderLeads(leadsData);
      renderCadence(cadenceData);
      renderObjections();

      // Chart 0: Contacted vs Goal Donut
      const ctxGoal = document.getElementById('goalChart').getContext('2d');
      new Chart(ctxGoal, {{
        type: 'doughnut',
        data: {{
          labels: ['Contacted (60)', 'Remaining Goal (20)'],
          datasets: [{{
            data: [60, 20],
            backgroundColor: ['#FF6B00', '#1F2937'],
            borderWidth: 0,
            cutout: '76%'
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ display: false }},
            tooltip: {{
              callbacks: {{
                label: (ctx) => ` ${{ctx.label}}`
              }}
            }}
          }}
        }}
      }});

      // Chart 1: Channel Conversion
      const ctxChan = document.getElementById('channelChart').getContext('2d');
      new Chart(ctxChan, {{
        type: 'bar',
        data: {{
          labels: ['Cold Email', 'LinkedIn', 'Cold Calling'],
          datasets: [
            {{
              label: 'Meeting Conversion Rate (%)',
              data: [8.0, 3.3, 3.6],
              backgroundColor: ['#2563EB', '#10B981', '#A855F7'],
              borderRadius: 8
            }}
          ]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ display: false }},
            tooltip: {{
              callbacks: {{
                label: (ctx) => ` ${{ctx.parsed.y}}% Meeting Conversion`
              }}
            }}
          }},
          scales: {{
            y: {{
              beginAtZero: true,
              max: 10,
              grid: {{ color: 'rgba(255, 255, 255, 0.05)' }},
              ticks: {{ color: '#94A3B8', callback: v => v + '%' }}
            }},
            x: {{
              grid: {{ display: false }},
              ticks: {{ color: '#94A3B8' }}
            }}
          }}
        }}
      }});

      // Chart 2: Objections Pie Chart
      const ctxObj = document.getElementById('objectionsChart').getContext('2d');
      new Chart(ctxObj, {{
        type: 'doughnut',
        data: {{
          labels: [
            'Using Postman/Jest',
            'Traffic Security/eBPF',
            'No Budget / Freeze',
            'Roadmap Priority'
          ],
          datasets: [{{
            data: [13, 10, 6, 5],
            backgroundColor: ['#FF6B00', '#F59E0B', '#3B82F6', '#8B5CF6'],
            borderColor: '#0F172A',
            borderWidth: 3
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{
              position: 'bottom',
              labels: {{ color: '#94A3B8', boxWidth: 12, font: {{ size: 10 }} }}
            }}
          }}
        }}
      }});

      // Chart 3: A/B Test Grouped Bar Chart
      const ctxAB = document.getElementById('abChart').getContext('2d');
      new Chart(ctxAB, {{
        type: 'bar',
        data: {{
          labels: ['Open Rate (%)', 'Reply Rate (%)', 'Meeting Conv (%)'],
          datasets: [
            {{
              label: 'Email A (Velocity)',
              data: [73.3, 26.7, 23.3],
              backgroundColor: '#10B981',
              borderRadius: 6
            }},
            {{
              label: 'Email B (Flaky Tests)',
              data: [53.3, 20.0, 13.3],
              backgroundColor: '#8B5CF6',
              borderRadius: 6
            }}
          ]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{
              position: 'top',
              labels: {{ color: '#94A3B8', boxWidth: 12, font: {{ size: 11 }} }}
            }}
          }},
          scales: {{
            y: {{
              beginAtZero: true,
              grid: {{ color: 'rgba(255, 255, 255, 0.05)' }},
              ticks: {{ color: '#94A3B8', callback: v => v + '%' }}
            }},
            x: {{
              grid: {{ display: false }},
              ticks: {{ color: '#94A3B8' }}
            }}
          }}
        }}
      }});

    }});
  </script>
</body>
</html>"""

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    print("Web dashboard generated at index.html")

# -------------------------------------------------------------
# MAIN EXECUTION
# -------------------------------------------------------------
if __name__ == "__main__":
    print("Generating DevTrace Sales Prospecting Assets...")
    cadence_logs = generate_cadence_logs(TARGET_COMPANIES)
    create_excel_workbook(TARGET_COMPANIES, cadence_logs)
    export_csv_files(TARGET_COMPANIES, cadence_logs)
    create_web_dashboard(TARGET_COMPANIES, cadence_logs)
    print("All artifacts generated successfully!")
