# B.C. local-government finance: moved

The B.C. model now has its own repository:
**[KushPatel29/bc-local-government-finance](https://github.com/KushPatel29/bc-local-government-finance)**.

It is a fuller and more careful build. It reads 90 official provincial files
(15 schedules, 2019–2024) through written contracts, rebuilds the numbers in a
tested dbt project, and holds the Excel model to an independent pandas
recomputation. It has 181 tests.

The version that used to sit here built its "asset condition ratio" on the
Province's published annual replacement-need field. That field is not usable
as published: Vancouver's reported need is 41% of its replacement value every
year. The new repository re-derives the renewal need instead and screens out
implausible values. That is why this copy was retired rather than kept
alongside it.
