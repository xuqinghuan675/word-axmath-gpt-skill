from __future__ import annotations

from normalize_axmath_tex import (
    canonical_prime_issues,
    contains_prime_or_derivative_marker,
    normalize_axmath_tex,
    prime_orders,
    limit_operator_placements,
    limit_operator_count,
    limit_operator_issues,
    smallint_issues,
)


def main() -> int:
    cases = {
        "$F′(x)$": r"$F\prime(x)$",
        "$F″(x)$": "$F''(x)$",
        "$F‴(x)$": "$F'''(x)$",
        "$F‴(ξ)$": "$F'''(ξ)$",
        "$y′²$": r"${y\prime}^{2}$",
        "$(y′)²$": r"$(y\prime)^{2}$",
        "$x_i′$": r"$x_i\prime$",
        r"$\frac{1+y′²}{y″}$": r"$\frac{1+{y\prime}^{2}}{y''}$",
        r"$f\prime(x)$": r"$f\prime(x)$",
        "$F'(x)$": r"$F\prime(x)$",
        "$F''(x)$": "$F''(x)$",
        "$F'''(x)$": "$F'''(x)$",
        r"${f}^{\prime2}$": r"${f\prime}^{2}$",
        r"${f}^{\prime 2}$": r"${f\prime}^{2}$",
        r"${f}^{''2}$": r"${f''}^{2}$",
        r"${f}^{'''2}$": r"${f'''}^{2}$",
        r"$\int {f}^{\prime2}(t)dt$": r"$\int {f\prime}^{2}(t)dt$",
    }
    for source, expected in cases.items():
        actual = normalize_axmath_tex(source)
        if actual != expected:
            raise AssertionError(f"{source!r} -> {actual!r}, expected {expected!r}")
        issues = canonical_prime_issues(actual)
        if issues:
            raise AssertionError(f"canonical output still has issues: {actual!r}: {issues}")

    limit_cases = {
        r"$\lim_{x\to0}f(x)$": r"$\underset{x\to0}{\lim}f(x)$",
        r"$\lim^{n\to\infty}a_n$": r"$\lim^{n\to\infty}a_n$",
        r"$\lim\nolimits_{x\to0}f(x)$": r"$\underset{x\to0}{\lim}f(x)$",
        r"$\lim\limits_{x\to0}f(x)$": r"$\underset{x\to0}{\lim}f(x)$",
        r"$\limsup_{n\to\infty}a_n$": r"$\limsup_{n\to\infty}a_n$",
        r"$\liminf_{n\to\infty}a_n$": r"$\liminf_{n\to\infty}a_n$",
    }
    for source, expected in limit_cases.items():
        actual = normalize_axmath_tex(source)
        if actual != expected:
            raise AssertionError(f"limit layout normalization {source!r} -> {actual!r}, expected {expected!r}")
        if limit_operator_issues(actual):
            raise AssertionError(f"normalized limit operator still violates AxMath placement: {actual!r}")
    if limit_operator_issues(r"$\lim_{x\to0}f(x)$") != ["lim_subscript_requires_underset"]:
        raise AssertionError("a lower lim script without underset must be rejected")
    if limit_operator_issues(r"$\lim\limits_{x\to0}f(x)$").count("lim_limits_command_not_allowed") != 1:
        raise AssertionError("lim\\\\limits must be rejected as unstable")
    if limit_operator_placements(r"$\underset{x\to0}{\lim}^{+}f(x)$") != [["below", "above"]]:
        raise AssertionError("a lim carrying both lower and upper scripts must preserve both placements")
    if limit_operator_count(r"$\underset{x\to0}{\lim}f(x)+\lim^{n\to\infty}a_n$") != 2:
        raise AssertionError("limit operator inventory must count both lower and upper scripts")
    if limit_operator_count(r"$\limsup_{n\to\infty}a_n$") != 0:
        raise AssertionError("limsup must not be misclassified as the lim operator")
    if normalize_axmath_tex(r"$\smallint f(x)$") != r"$\int f(x)$":
        raise AssertionError("smallint must normalize to the required integral form")
    if smallint_issues(r"$\smallint f(x)$") != ["smallint_forbidden"]:
        raise AssertionError("raw AxMath output must reject smallint")

    order_cases = {
        r"$F\prime(x)$": [1],
        "$F''(x)$": [2],
        "$F'''(x)$": [3],
        r"$\frac{F\prime(x)+G''(x)}{H'''(x)}$": [1, 2, 3],
        r"${y\prime}^{2}$": [1],
    }
    for tex, expected_orders in order_cases.items():
        actual_orders = prime_orders(tex)
        if actual_orders != expected_orders:
            raise AssertionError(
                f"prime_orders({tex!r}) -> {actual_orders!r}, expected {expected_orders!r}"
            )

    for source in ["y′", "y″", "y‴", "x'", r"f\prime(x)", "x_i’"]:
        if not contains_prime_or_derivative_marker(source):
            raise AssertionError(f"prime detector missed {source!r}")

    if canonical_prime_issues("$y′$") != ["noncanonical_prime_literal"]:
        raise AssertionError("Unicode prime literal must be rejected by the canonical contract")
    if "noncanonical_single_prime_apostrophe" not in canonical_prime_issues("$y'$"):
        raise AssertionError(r"single ASCII apostrophe must normalize to AxMath canonical \prime")
    if "unverified_prime_order" not in canonical_prime_issues("$y''''$"):
        raise AssertionError("prime orders above 3 must not be guessed")
    if "ungrouped_primed_atom_exponent" not in canonical_prime_issues(r"$y\prime^{2}$"):
        raise AssertionError(r"y\prime^{2} must be grouped as {y\prime}^{2}")
    if "ungrouped_primed_atom_exponent" not in canonical_prime_issues("$y''^{2}$"):
        raise AssertionError("y''^{2} must be grouped as {y''}^{2}")
    if "fused_prime_exponent" not in canonical_prime_issues(r"${f}^{\prime2}$"):
        raise AssertionError(r"raw OMML-style {f}^{\prime2} must be rejected until normalized")

    print("PRIME_NORMALIZATION_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
