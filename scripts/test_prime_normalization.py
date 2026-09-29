from __future__ import annotations

from normalize_axmath_tex import (
    canonical_prime_issues,
    contains_prime_or_derivative_marker,
    normalize_axmath_tex,
    prime_orders,
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
