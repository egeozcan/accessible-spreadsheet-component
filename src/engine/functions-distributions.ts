/** Regression and probability-distribution built-ins (values only). */
import { type Registrar } from './functions.js';
import { flattenDeep, isErrorCode, toBoolean, toNumber } from './values.js';

const SQRT_2PI = Math.sqrt(2 * Math.PI);
const SQRT_2 = Math.SQRT2;
const FPMIN = 1e-300;
const EPS = 1e-15;
const MAX_ITER = 100000;

/** Throws an Excel error code. */
function fail(code = '#NUM!'): never {
  throw new Error(code);
}

/** The value itself when finite, otherwise #NUM!. */
function finite(v: number): number {
  return Number.isFinite(v) ? v : fail();
}

/** An optional numeric argument (undefined when omitted). */
function optional(v: unknown, fallback: number): number {
  return v === undefined ? fallback : toNumber(v);
}

// ─── Special functions ──────────────────────────────

/**
 * ln Γ(x) for x > 0: shift x up to 15 with the recurrence, then the Stirling
 * series (Bernoulli terms through x^-13), which is accurate to about 1e-16 there.
 */
function lnGamma(x: number): number {
  let shifted = x;
  let product = 1;
  while (shifted < 15) {
    product *= shifted;
    shifted += 1;
  }
  const z = 1 / (shifted * shifted);
  const series =
    (1 / shifted) *
    (1 / 12 - z * (1 / 360 - z * (1 / 1260 - z * (1 / 1680 - z * (1 / 1188 - z * (691 / 360360 - z / 156))))));
  return (shifted - 0.5) * Math.log(shifted) - shifted + 0.5 * Math.log(2 * Math.PI) + series - Math.log(product);
}

/** Γ(x) for any real x that is not a non-positive integer. */
function gamma(x: number): number {
  if (Number.isInteger(x) && x > 0 && x <= 171) {
    let product = 1;
    for (let i = 2; i < x; i++) product *= i;
    return product;
  }
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
  return Math.exp(lnGamma(x));
}

function lnBeta(a: number, b: number): number {
  return lnGamma(a) + lnGamma(b) - lnGamma(a + b);
}

/** Lower regularized incomplete gamma P(a, x) by its power series (x < a + 1). */
function gammaSeries(a: number, x: number): number {
  let ap = a;
  let del = 1 / a;
  let sum = del;
  for (let n = 0; n < MAX_ITER; n++) {
    ap += 1;
    del *= x / ap;
    sum += del;
    if (Math.abs(del) < Math.abs(sum) * EPS) break;
  }
  return sum * Math.exp(-x + a * Math.log(x) - lnGamma(a));
}

/** Upper regularized incomplete gamma Q(a, x) by its continued fraction (x >= a + 1). */
function gammaContinued(a: number, x: number): number {
  let b = x + 1 - a;
  let c = 1 / FPMIN;
  let d = 1 / b;
  let h = d;
  for (let i = 1; i < MAX_ITER; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = b + an / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return Math.exp(-x + a * Math.log(x) - lnGamma(a)) * h;
}

/** Regularized lower incomplete gamma P(a, x). */
function gammaP(a: number, x: number): number {
  if (x <= 0) return 0;
  return x < a + 1 ? gammaSeries(a, x) : 1 - gammaContinued(a, x);
}

/** Regularized upper incomplete gamma Q(a, x). */
function gammaQ(a: number, x: number): number {
  if (x <= 0) return 1;
  return x < a + 1 ? 1 - gammaSeries(a, x) : gammaContinued(a, x);
}

/** Continued fraction for the incomplete beta function (modified Lentz). */
function betaContinued(x: number, a: number, b: number): number {
  const qab = a + b;
  const qap = a + 1;
  const qam = a - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d;
  let h = d;
  for (let m = 1; m < MAX_ITER; m++) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return h;
}

/** Regularized incomplete beta I_x(a, b) for 0 <= x <= 1. */
function betaInc(x: number, a: number, b: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const front = Math.exp(a * Math.log(x) + b * Math.log1p(-x) - lnBeta(a, b));
  return x < (a + 1) / (a + b + 2)
    ? (front * betaContinued(x, a, b)) / a
    : 1 - (front * betaContinued(1 - x, b, a)) / b;
}

function erf(x: number): number {
  return x < 0 ? -gammaP(0.5, x * x) : gammaP(0.5, x * x);
}

function erfc(x: number): number {
  return x < 0 ? 1 + gammaP(0.5, x * x) : gammaQ(0.5, x * x);
}

function pdfNormal(z: number): number {
  return Math.exp((-z * z) / 2) / SQRT_2PI;
}

function cdfNormal(z: number): number {
  return 0.5 * erfc(-z / SQRT_2);
}

// Acklam's rational approximation to the normal quantile for p <= 0.5 (relative error ~1e-9),
// refined by Halley steps against the exact CDF.
const AK_A = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
const AK_B = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
const AK_C = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
const AK_D = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];

/** Standard normal quantile for 0 < p < 1. */
function normInv(p: number): number {
  if (p > 0.5) return -normInv(1 - p);
  let x: number;
  if (p < 0.02425) {
    const q = Math.sqrt(-2 * Math.log(p));
    const [c0, c1, c2, c3, c4, c5] = AK_C;
    const [d0, d1, d2, d3] = AK_D;
    x = (((((c0 * q + c1) * q + c2) * q + c3) * q + c4) * q + c5) / ((((d0 * q + d1) * q + d2) * q + d3) * q + 1);
  } else {
    const q = p - 0.5;
    const r = q * q;
    const [a0, a1, a2, a3, a4, a5] = AK_A;
    const [b0, b1, b2, b3, b4] = AK_B;
    x = ((((((a0 * r + a1) * r + a2) * r + a3) * r + a4) * r + a5) * q) /
      (((((b0 * r + b1) * r + b2) * r + b3) * r + b4) * r + 1);
  }
  for (let i = 0; i < 3; i++) {
    const e = cdfNormal(x) - p;
    const u = (e * SQRT_2PI) * Math.exp((x * x) / 2);
    x -= u / (1 + (x * u) / 2);
  }
  return x;
}

/** ln C(n, k) for integers 0 <= k <= n (-Infinity outside). */
function lnChoose(n: number, k: number): number {
  if (k < 0 || k > n) return -Infinity;
  const m = Math.min(k, n - k);
  // Huge counts: through log-gamma, so the work is bounded
  if (m > 100_000) return lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1);
  let s = 0;
  for (let i = 0; i < m; i++) s += Math.log(n - i) - Math.log(i + 1);
  return s;
}

/** Smallest x >= 0 with cdf(x) = p, for an increasing cdf on [0, inf). */
function invertUp(cdf: (x: number) => number, p: number): number {
  let lo = 0;
  let hi = 1;
  while (cdf(hi) < p && hi < 1e300) hi *= 2;
  for (let i = 0; i < 2200; i++) {
    const mid = (lo + hi) / 2;
    if (mid <= lo || mid >= hi) break;
    if (cdf(mid) < p) lo = mid;
    else hi = mid;
  }
  return (lo + hi) / 2;
}

/** Smallest x >= 0 with sf(x) = p, for a decreasing survival function on [0, inf). */
function invertTail(sf: (x: number) => number, p: number): number {
  let lo = 0;
  let hi = 1;
  while (sf(hi) > p && hi < 1e300) hi *= 2;
  for (let i = 0; i < 2200; i++) {
    const mid = (lo + hi) / 2;
    if (mid <= lo || mid >= hi) break;
    if (sf(mid) > p) lo = mid;
    else hi = mid;
  }
  return (lo + hi) / 2;
}

/** Upper tail P(T > t) of Student's t with df degrees of freedom (df > 0). */
function tailT(t: number, df: number): number {
  const half = 0.5 * betaInc(df / (df + t * t), df / 2, 0.5);
  return t >= 0 ? half : 1 - half;
}

/** Density of a gamma distribution with shape a and scale b. */
function gammaPdf(x: number, a: number, b: number): number {
  if (x < 0) return 0;
  if (x === 0) {
    if (a < 1) return fail();
    return a === 1 ? 1 / b : 0;
  }
  return Math.exp((a - 1) * Math.log(x / b) - x / b - lnGamma(a)) / b;
}

/** Binomial probability of k successes in n trials. */
function binomPmf(k: number, n: number, p: number): number {
  if (p === 0) return k === 0 ? 1 : 0;
  if (p === 1) return k === n ? 1 : 0;
  return Math.exp(lnChoose(n, k) + k * Math.log(p) + (n - k) * Math.log1p(-p));
}

/** Gamma-like shape check shared by the chi-square functions. */
function degreesOfFreedom(df: number): number {
  const d = Math.trunc(df);
  if (d < 1 || d >= 1e10) fail();
  return d;
}

// ─── Registration ──────────────────────────────────

const arity = (min: number, max = min) => ({ arity: [min, max] as [number, number] });

export function registerDistributionsBuiltins(reg: Registrar): void {
  const { def } = reg;

  // Normal
  const normDist = (_ctx: unknown, x: unknown, mean: unknown, sd: unknown, cum: unknown) => {
    const s = toNumber(sd);
    if (s <= 0) fail();
    const z = (toNumber(x) - toNumber(mean)) / s;
    return finite(toBoolean(cum) ? cdfNormal(z) : pdfNormal(z) / s);
  };
  for (const name of ['NORM.DIST', 'NORMDIST']) def(name, 'nnnn', normDist, arity(4));

  const normSDist = (_ctx: unknown, z: unknown, cum: unknown) => {
    const v = toNumber(z);
    return finite(toBoolean(cum) ? cdfNormal(v) : pdfNormal(v));
  };
  def('NORM.S.DIST', 'nn', normSDist, arity(2));

  def('NORMSDIST', 'n', (_ctx, z) => finite(cdfNormal(toNumber(z))), arity(1));

  const normInvFn = (_ctx: unknown, p: unknown, mean: unknown, sd: unknown) => {
    const pv = toNumber(p);
    const s = toNumber(sd);
    if (pv <= 0 || pv >= 1 || s <= 0) fail();
    return finite(toNumber(mean) + s * normInv(pv));
  };
  for (const name of ['NORM.INV', 'NORMINV']) def(name, 'nnn', normInvFn, arity(3));

  const normSInv = (_ctx: unknown, p: unknown) => {
    const pv = toNumber(p);
    if (pv <= 0 || pv >= 1) fail();
    return finite(normInv(pv));
  };
  for (const name of ['NORM.S.INV', 'NORMSINV']) def(name, 'n', normSInv, arity(1));

  def('GAUSS', 'n', (_ctx, z) => finite(cdfNormal(toNumber(z)) - 0.5), arity(1));
  def('PHI', 'n', (_ctx, x) => finite(pdfNormal(toNumber(x))), arity(1));

  // Log-normal
  def(
    'LOGNORM.DIST',
    'nnnn',
    (_ctx, x, mean, sd, cum) => {
      const v = toNumber(x);
      const s = toNumber(sd);
      if (v <= 0 || s <= 0) fail();
      const z = (Math.log(v) - toNumber(mean)) / s;
      return finite(toBoolean(cum) ? cdfNormal(z) : pdfNormal(z) / (v * s));
    },
    arity(4)
  );

  def(
    'LOGNORM.INV',
    'nnn',
    (_ctx, p, mean, sd) => {
      const pv = toNumber(p);
      const s = toNumber(sd);
      if (pv <= 0 || pv >= 1 || s <= 0) fail();
      return finite(Math.exp(toNumber(mean) + s * normInv(pv)));
    },
    arity(3)
  );

  // Binomial
  const binomDist = (_ctx: unknown, k: unknown, n: unknown, p: unknown, cum: unknown) => {
    const kv = Math.trunc(toNumber(k));
    const nv = Math.trunc(toNumber(n));
    const pv = toNumber(p);
    if (kv < 0 || nv < 0 || kv > nv || pv < 0 || pv > 1) fail();
    if (!toBoolean(cum)) return binomPmf(kv, nv, pv);
    return kv === nv ? 1 : betaInc(1 - pv, nv - kv, kv + 1);
  };
  for (const name of ['BINOM.DIST', 'BINOMDIST']) def(name, 'nnnn', binomDist, arity(4));

  def(
    'BINOM.INV',
    'nnn',
    (_ctx, n, p, alpha) => {
      const nv = Math.trunc(toNumber(n));
      const pv = toNumber(p);
      const av = toNumber(alpha);
      // Excel (checked against it) also rejects an alpha of exactly 0 or 1
      if (nv < 0 || pv < 0 || pv > 1 || av <= 0 || av >= 1) fail();
      // Smallest k whose cumulative probability reaches alpha: binary search on the cdf
      // (a hair of tolerance, so an exact 0.5 at the median is not missed by rounding)
      if (!Number.isSafeInteger(nv)) fail();
      const cdf = (k: number) => (k >= nv ? 1 : betaInc(1 - pv, nv - k, k + 1));
      let lo = 0;
      let hi = nv;
      while (lo < hi) {
        const mid = lo + Math.floor((hi - lo) / 2); // safe near 2^53
        if (cdf(mid) >= av * (1 - 1e-14)) hi = mid; // smallest k with cdf >= alpha (Excel's own exact ties are float noise: n=7 steps past one, n=9 does not)
        else lo = mid + 1;
      }
      return lo;
    },
    arity(3)
  );

  // Poisson
  const poisson = (_ctx: unknown, x: unknown, mean: unknown, cum: unknown) => {
    const k = Math.trunc(toNumber(x));
    const m = toNumber(mean);
    if (k < 0 || m < 0) fail();
    if (toBoolean(cum)) return gammaQ(k + 1, m);
    if (m === 0) return k === 0 ? 1 : 0;
    return Math.exp(k * Math.log(m) - m - lnGamma(k + 1));
  };
  for (const name of ['POISSON.DIST', 'POISSON']) def(name, 'nnn', poisson, arity(3));

  // Exponential
  const expon = (_ctx: unknown, x: unknown, lambda: unknown, cum: unknown) => {
    const v = toNumber(x);
    const l = toNumber(lambda);
    if (v < 0 || l <= 0) fail();
    return finite(toBoolean(cum) ? -Math.expm1(-l * v) : l * Math.exp(-l * v));
  };
  for (const name of ['EXPON.DIST', 'EXPONDIST']) def(name, 'nnn', expon, arity(3));

  // Gamma family
  def(
    'GAMMA',
    'n',
    (_ctx, x) => {
      const v = toNumber(x);
      if (v === 0 || (v < 0 && Number.isInteger(v))) fail();
      return finite(gamma(v));
    },
    arity(1)
  );

  for (const name of ['GAMMALN', 'GAMMALN.PRECISE']) {
    def(
      name,
      'n',
      (_ctx, x) => {
        const v = toNumber(x);
        if (v <= 0) fail();
        return finite(lnGamma(v));
      },
      arity(1)
    );
  }

  def(
    'GAMMA.DIST',
    'nnnn',
    (_ctx, x, alpha, beta, cum) => {
      const v = toNumber(x);
      const a = toNumber(alpha);
      const b = toNumber(beta);
      if (v < 0 || a <= 0 || b <= 0) fail();
      return finite(toBoolean(cum) ? gammaP(a, v / b) : gammaPdf(v, a, b));
    },
    arity(4)
  );

  def(
    'GAMMA.INV',
    'nnn',
    (_ctx, p, alpha, beta) => {
      const pv = toNumber(p);
      const a = toNumber(alpha);
      const b = toNumber(beta);
      if (pv < 0 || pv > 1 || a <= 0 || b <= 0 || pv === 1) fail();
      if (pv === 0) return 0;
      return finite(b * invertUp((x) => gammaP(a, x), pv));
    },
    arity(3)
  );

  // Beta
  def(
    'BETA.DIST',
    'nnnnnn',
    (_ctx, x, alpha, beta, cum, lower, upper) => {
      const v = toNumber(x);
      const a = toNumber(alpha);
      const b = toNumber(beta);
      const A = optional(lower, 0);
      const B = optional(upper, 1);
      if (a <= 0 || b <= 0 || B <= A || v < A || v > B) fail();
      const z = (v - A) / (B - A);
      if (toBoolean(cum)) return finite(betaInc(z, a, b));
      if ((z === 0 && a < 1) || (z === 1 && b < 1)) fail();
      const logTerm = (e: number, t: number) => (e === 0 ? 0 : e * Math.log(t));
      const density =
        Math.exp(logTerm(a - 1, z) + logTerm(b - 1, 1 - z) - lnBeta(a, b)) / (B - A);
      return finite(density);
    },
    arity(4, 6)
  );

  def(
    'BETA.INV',
    'nnnnn',
    (_ctx, p, alpha, beta, lower, upper) => {
      const pv = toNumber(p);
      const a = toNumber(alpha);
      const b = toNumber(beta);
      const A = optional(lower, 0);
      const B = optional(upper, 1);
      if (pv < 0 || pv > 1 || a <= 0 || b <= 0 || B <= A) fail();
      return finite(A + (B - A) * invertUp((x) => betaInc(Math.min(x, 1), a, b), pv));
    },
    arity(3, 5)
  );

  // Chi-square
  def(
    'CHISQ.DIST',
    'nnn',
    (_ctx, x, df, cum) => {
      const v = toNumber(x);
      const k = degreesOfFreedom(toNumber(df));
      if (v < 0) fail();
      return finite(toBoolean(cum) ? gammaP(k / 2, v / 2) : gammaPdf(v, k / 2, 2));
    },
    arity(3)
  );

  def(
    'CHISQ.DIST.RT',
    'nn',
    (_ctx, x, df) => {
      const v = toNumber(x);
      const k = degreesOfFreedom(toNumber(df));
      if (v < 0) fail();
      return finite(gammaQ(k / 2, v / 2));
    },
    arity(2)
  );

  def(
    'CHISQ.INV',
    'nn',
    (_ctx, p, df) => {
      const pv = toNumber(p);
      const k = degreesOfFreedom(toNumber(df));
      if (pv < 0 || pv >= 1) fail();
      return finite(2 * invertUp((x) => gammaP(k / 2, x), pv));
    },
    arity(2)
  );

  def(
    'CHISQ.INV.RT',
    'nn',
    (_ctx, p, df) => {
      const pv = toNumber(p);
      const k = degreesOfFreedom(toNumber(df));
      if (pv <= 0 || pv > 1) fail();
      return finite(2 * invertTail((x) => gammaQ(k / 2, x), pv));
    },
    arity(2)
  );

  // Student's t
  def(
    'T.DIST',
    'nnn',
    (_ctx, x, df, cum) => {
      const v = toNumber(x);
      const k = Math.trunc(toNumber(df));
      if (k < 1) fail();
      if (toBoolean(cum)) return finite(tailT(-v, k));
      const density = Math.exp(
        lnGamma((k + 1) / 2) - lnGamma(k / 2) - 0.5 * Math.log(k * Math.PI) -
          ((k + 1) / 2) * Math.log1p((v * v) / k)
      );
      return finite(density);
    },
    arity(3)
  );

  def(
    'T.DIST.RT',
    'nn',
    (_ctx, x, df) => {
      const k = Math.trunc(toNumber(df));
      if (k < 1) fail();
      return finite(tailT(toNumber(x), k));
    },
    arity(2)
  );

  def(
    'T.DIST.2T',
    'nn',
    (_ctx, x, df) => {
      const v = toNumber(x);
      const k = Math.trunc(toNumber(df));
      if (v < 0 || k < 1) fail();
      return finite(2 * tailT(v, k));
    },
    arity(2)
  );

  def(
    'T.INV',
    'nn',
    (_ctx, p, df) => {
      const pv = toNumber(p);
      const k = Math.trunc(toNumber(df));
      if (pv <= 0 || pv >= 1 || k < 1) fail();
      if (pv === 0.5) return 0;
      // Lower tail: -t is the upper-tail quantile of 1 - p, found on the positive side.
      return finite(
        pv < 0.5
          ? -invertTail((x) => tailT(x, k), pv)
          : invertTail((x) => tailT(x, k), 1 - pv)
      );
    },
    arity(2)
  );

  def(
    'T.INV.2T',
    'nn',
    (_ctx, p, df) => {
      const pv = toNumber(p);
      const k = Math.trunc(toNumber(df));
      if (pv <= 0 || pv > 1 || k < 1) fail();
      return finite(invertTail((x) => 2 * tailT(x, k), pv));
    },
    arity(2)
  );

  // F
  const fDensity = (x: number, d1: number, d2: number) => {
    if (x === 0) {
      if (d1 < 2) fail();
      return d1 === 2 ? 1 : 0;
    }
    const logB = lnBeta(d1 / 2, d2 / 2);
    return Math.exp(
      (d1 / 2) * Math.log(d1 / d2) + (d1 / 2 - 1) * Math.log(x) - ((d1 + d2) / 2) * Math.log1p((d1 * x) / d2) - logB
    );
  };
  const fCdf = (x: number, d1: number, d2: number) => (x <= 0 ? 0 : betaInc((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2));
  const fTail = (x: number, d1: number, d2: number) => (x <= 0 ? 1 : betaInc(d2 / (d2 + d1 * x), d2 / 2, d1 / 2));

  def(
    'F.DIST',
    'nnnn',
    (_ctx, x, df1, df2, cum) => {
      const v = toNumber(x);
      const d1 = degreesOfFreedom(toNumber(df1));
      const d2 = degreesOfFreedom(toNumber(df2));
      if (v < 0) fail();
      return finite(toBoolean(cum) ? fCdf(v, d1, d2) : fDensity(v, d1, d2));
    },
    arity(4)
  );

  def(
    'F.DIST.RT',
    'nnn',
    (_ctx, x, df1, df2) => {
      const v = toNumber(x);
      const d1 = degreesOfFreedom(toNumber(df1));
      const d2 = degreesOfFreedom(toNumber(df2));
      if (v < 0) fail();
      return finite(fTail(v, d1, d2));
    },
    arity(3)
  );

  def(
    'F.INV',
    'nnn',
    (_ctx, p, df1, df2) => {
      const pv = toNumber(p);
      const d1 = degreesOfFreedom(toNumber(df1));
      const d2 = degreesOfFreedom(toNumber(df2));
      if (pv < 0 || pv >= 1) fail();
      return finite(invertUp((x) => fCdf(x, d1, d2), pv));
    },
    arity(3)
  );

  def(
    'F.INV.RT',
    'nnn',
    (_ctx, p, df1, df2) => {
      const pv = toNumber(p);
      const d1 = degreesOfFreedom(toNumber(df1));
      const d2 = degreesOfFreedom(toNumber(df2));
      if (pv <= 0 || pv > 1) fail();
      return finite(invertTail((x) => fTail(x, d1, d2), pv));
    },
    arity(3)
  );

  // Weibull
  def(
    'WEIBULL.DIST',
    'nnnn',
    (_ctx, x, alpha, beta, cum) => {
      const v = toNumber(x);
      const a = toNumber(alpha);
      const b = toNumber(beta);
      if (v < 0 || a <= 0 || b <= 0) fail();
      if (toBoolean(cum)) return finite(-Math.expm1(-Math.pow(v / b, a)));
      if (v === 0) {
        if (a < 1) fail();
        return a === 1 ? 1 / b : 0;
      }
      const t = Math.pow(v / b, a - 1);
      return finite((a / b) * t * Math.exp(-Math.pow(v / b, a)));
    },
    arity(4)
  );

  // Hypergeometric: sample_s successes in number_sample draws from a population of
  // population_s successes among number_pop.
  def(
    'HYPGEOM.DIST',
    'nnnnn',
    (_ctx, sampleS, numberSample, popS, numberPop, cum) => {
      const s = Math.trunc(toNumber(sampleS));
      const n = Math.trunc(toNumber(numberSample));
      const K = Math.trunc(toNumber(popS));
      const N = Math.trunc(toNumber(numberPop));
      if (s < 0 || n < 0 || K < 0 || N < 0 || s > n || s > K || n - s > N - K || n > N || K > N) fail();
      const lowest = Math.max(0, n - (N - K));
      if (!Number.isSafeInteger(N) || s - lowest > 5_000_000) fail('#NUM!'); // bounded work
      const pmf = (i: number) =>
        Math.exp(lnChoose(K, i) + lnChoose(N - K, n - i) - lnChoose(N, n));
      if (!toBoolean(cum)) return finite(pmf(s));
      // Walk away from the mode with the term ratio, summing whichever tail is shorter:
      // the work is about the width of the distribution, not the size of the counts
      const mode = Math.floor(((n + 1) * (K + 1)) / (N + 2));
      const highest = Math.min(n, K);
      const walk = (from: number, step: 1 | -1, last: number): number => {
        let term = pmf(from);
        let sum = 0;
        for (let i = from, k = 0; step === 1 ? i <= last : i >= last; i += step, k++) {
          if (k > 5_000_000) fail('#NUM!');
          sum += term;
          // ratio of the next term (i + step) to this one
          term *=
            step === 1
              ? ((K - i) * (n - i)) / ((i + 1) * (N - K - n + i + 1))
              : (i * (N - K - n + i)) / ((K - i + 1) * (n - i + 1));
          if (!(term > sum * 1e-18) && (step === 1 ? i >= mode : i <= mode)) break;
        }
        return sum;
      };
      if (s >= highest) return 1;
      if (s <= mode) return finite(Math.min(1, walk(s, -1, lowest)));
      return finite(Math.max(0, 1 - walk(s + 1, 1, highest)));
    },
    arity(5)
  );

  // Negative binomial: number_f failures before number_s successes.
  def(
    'NEGBINOM.DIST',
    'nnnn',
    (_ctx, numberF, numberS, probability, cum) => {
      const f = Math.trunc(toNumber(numberF));
      const s = Math.trunc(toNumber(numberS));
      const p = toNumber(probability);
      if (f < 0 || s < 1 || p < 0 || p > 1) fail();
      if (toBoolean(cum)) return betaInc(p, s, f + 1);
      const logTerm = f === 0 ? 0 : f * Math.log1p(-p);
      const sTerm = p === 0 ? -Infinity : s * Math.log(p);
      return finite(Math.exp(lnChoose(f + s - 1, f) + sTerm + logTerm));
    },
    arity(4)
  );

  // Confidence intervals (half-widths)
  const confidenceNorm = (_ctx: unknown, alpha: unknown, sd: unknown, size: unknown) => {
    const a = toNumber(alpha);
    const s = toNumber(sd);
    const n = Math.trunc(toNumber(size));
    if (a <= 0 || a >= 1 || s <= 0 || n < 1) fail();
    return finite(-normInv(a / 2) * (s / Math.sqrt(n)));
  };
  for (const name of ['CONFIDENCE.NORM', 'CONFIDENCE']) def(name, 'nnn', confidenceNorm, arity(3));

  def(
    'CONFIDENCE.T',
    'nnn',
    (_ctx, alpha, sd, size) => {
      const a = toNumber(alpha);
      const s = toNumber(sd);
      const n = Math.trunc(toNumber(size));
      if (a <= 0 || a >= 1 || s <= 0 || n < 1) fail();
      if (n === 1) fail('#DIV/0!');
      const df = n - 1;
      const t = invertTail((x) => 2 * tailT(x, df), a);
      return finite(t * (s / Math.sqrt(n)));
    },
    arity(3)
  );

  // One-tailed z-test: P(Z > z) for the mean of the data against x.
  def(
    'Z.TEST',
    'bnn',
    (_ctx, array, x, sigma) => {
      const data: number[] = [];
      for (const v of flattenDeep(array)) {
        if (isErrorCode(v)) fail(v);
        if (typeof v === 'number') data.push(v);
      }
      const n = data.length;
      if (n === 0) fail('#N/A');
      const mean = data.reduce((sum, v) => sum + v, 0) / n;
      let sd: number;
      if (sigma !== undefined) {
        sd = toNumber(sigma);
      } else {
        if (n < 2) fail('#DIV/0!');
        sd = Math.sqrt(data.reduce((sum, v) => sum + (v - mean) ** 2, 0) / (n - 1));
      }
      if (sd <= 0) fail('#NUM!');
      const z = (mean - toNumber(x)) / (sd / Math.sqrt(n));
      return finite(cdfNormal(-z));
    },
    arity(2, 3)
  );

  // Error function family
  def(
    'ERF',
    'nn',
    (_ctx, lower, upper) =>
      finite(upper === undefined ? erf(toNumber(lower)) : erf(toNumber(upper)) - erf(toNumber(lower))),
    arity(1, 2)
  );
  def('ERF.PRECISE', 'n', (_ctx, x) => finite(erf(toNumber(x))), arity(1));
  for (const name of ['ERFC', 'ERFC.PRECISE']) def(name, 'n', (_ctx, x) => finite(erfc(toNumber(x))), arity(1));
}
