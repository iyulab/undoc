use super::*;

fn shape(name: &str, w: f64, h: f64) -> Geometry {
    evaluate(preset(name).expect(name), w, h, &[])
}

/// The points a path visits (the end point of every segment), rounded to 1e-6.
fn points(outline: &Outline) -> Vec<(f64, f64)> {
    let r = |v: f64| (v * 1e6).round() / 1e6;
    outline
        .segments
        .iter()
        .filter_map(|s| match *s {
            Segment::Move(p) | Segment::Line(p) | Segment::Quad(_, p) | Segment::Cubic(_, _, p) => {
                Some((r(p.x), r(p.y)))
            }
            Segment::Close => None,
        })
        .collect()
}

#[test]
fn every_preset_is_found_by_name_and_the_table_is_sorted() {
    // 187 definitions in the annex, upDownArrow among them twice.
    assert_eq!(presets::PRESETS.len(), 186);
    assert!(presets::PRESETS.windows(2).all(|w| w[0].name < w[1].name));
    for p in presets::PRESETS {
        assert!(std::ptr::eq(preset(p.name).unwrap(), p));
    }
    assert!(preset("noSuchShape").is_none());
}

#[test]
fn a_rectangle_is_its_four_corners() {
    let g = shape("rect", 200.0, 100.0);
    assert_eq!(
        points(&g.outlines[0]),
        [(0.0, 0.0), (200.0, 0.0), (200.0, 100.0), (0.0, 100.0)]
    );
    assert_eq!(g.text_rect, (0.0, 0.0, 200.0, 100.0));
}

/// rightArrow at its default adjust values (50000, 50000) on a 200 × 100 shape: the shaft is
/// half the height, the head as long as half the short side.
#[test]
fn a_right_arrow_follows_its_guide_formulas() {
    let g = shape("rightArrow", 200.0, 100.0);
    assert_eq!(
        points(&g.outlines[0]),
        [
            (0.0, 25.0),
            (150.0, 25.0),
            (150.0, 0.0),
            (200.0, 50.0),
            (150.0, 100.0),
            (150.0, 75.0),
            (0.0, 75.0)
        ]
    );
}

#[test]
fn an_adjust_value_overrides_the_default() {
    // The shaft as thick as the whole height.
    let g = evaluate(
        preset("rightArrow").unwrap(),
        200.0,
        100.0,
        &[("adj1", 100_000.0)],
    );
    assert_eq!(points(&g.outlines[0])[0], (0.0, 0.0));
}

/// An ellipse is four quarter arcs through the midpoints of its bounding box.
#[test]
fn an_ellipse_passes_through_the_edges_of_its_box() {
    let g = shape("ellipse", 200.0, 100.0);
    let pts = points(&g.outlines[0]);
    for edge in [(0.0, 50.0), (100.0, 0.0), (200.0, 50.0), (100.0, 100.0)] {
        assert!(pts.contains(&edge), "{edge:?} not in {pts:?}");
    }
    // Every control point stays within the box (an arc bulging outward would leave it).
    for s in &g.outlines[0].segments {
        if let Segment::Cubic(c1, c2, _) = s {
            for c in [c1, c2] {
                assert!(
                    (-1e-9..=200.0 + 1e-9).contains(&c.x) && (-1e-9..=100.0 + 1e-9).contains(&c.y)
                );
            }
        }
    }
}

/// A rounded rectangle's corner arcs start and end where its straight edges do.
#[test]
fn a_rounded_rectangle_joins_its_arcs_to_its_edges() {
    // Default adj 16667 of the short side: a radius of 100 · 0.16667 ≈ 16.667.
    let g = shape("roundRect", 200.0, 100.0);
    let pts = points(&g.outlines[0]);
    let r = (100.0 * 16_667.0 / 100_000.0 * 1e6_f64).round() / 1e6;
    assert_eq!(pts[0], (0.0, r));
    assert!(pts.contains(&(r, 0.0)));
    assert!(pts.contains(&(200.0 - r, 0.0)));
    assert!(pts.contains(&(200.0, 100.0 - r)));
}

/// A path with its own coordinate size is stretched over the shape.
#[test]
fn a_sized_path_is_stretched_over_the_shape() {
    // flowChartProcess draws a 1 × 1 path.
    let def = preset("flowChartProcess").unwrap();
    assert_eq!((def.paths[0].w, def.paths[0].h), (Some(1), Some(1)));
    let g = evaluate(def, 300.0, 120.0, &[]);
    assert_eq!(
        points(&g.outlines[0]),
        [(0.0, 0.0), (300.0, 0.0), (300.0, 120.0), (0.0, 120.0)]
    );
}

/// Every preset evaluates to finite coordinates at every aspect, including a degenerate one,
/// and names only guides it defines — an unknown name would read as 0 and pull a point to
/// the corner without any error.
#[test]
fn every_preset_evaluates_to_finite_points_with_only_known_guides() {
    let builtins = Env::new(1.0, 1.0);
    for p in presets::PRESETS {
        let mut known: Vec<&str> = builtins.values.keys().copied().collect();
        for g in p.av.iter().chain(p.gd) {
            for token in g.fmla.split_whitespace().skip(1) {
                assert!(
                    token.parse::<f64>().is_ok() || known.contains(&token),
                    "{}: {} names unknown guide {token:?}",
                    p.name,
                    g.name
                );
            }
            known.push(g.name);
        }
        let check = |token: &str| {
            assert!(
                token.parse::<f64>().is_ok() || known.contains(&token),
                "{}: path names unknown guide {token:?}",
                p.name
            )
        };
        for path in p.paths {
            for cmd in path.cmds {
                match *cmd {
                    Cmd::Move(x, y) | Cmd::Line(x, y) => [x, y].into_iter().for_each(check),
                    Cmd::Arc { wr, hr, st, sw } => [wr, hr, st, sw].into_iter().for_each(check),
                    Cmd::Quad(a) => a.into_iter().for_each(check),
                    Cmd::Cubic(a) => a.into_iter().for_each(check),
                    Cmd::Close => {}
                }
            }
        }
        for (w, h) in [(100.0, 100.0), (400.0, 100.0), (100.0, 400.0), (1.0, 0.0)] {
            let g = evaluate(p, w, h, &[]);
            assert!(!g.outlines.is_empty(), "{} has no path", p.name);
            for o in &g.outlines {
                for s in &o.segments {
                    let pts: Vec<Point> = match *s {
                        Segment::Move(a) | Segment::Line(a) => vec![a],
                        Segment::Quad(a, b) => vec![a, b],
                        Segment::Cubic(a, b, c) => vec![a, b, c],
                        Segment::Close => vec![],
                    };
                    for pt in pts {
                        assert!(
                            pt.x.is_finite() && pt.y.is_finite(),
                            "{} at {w}x{h}: {pt:?}",
                            p.name
                        );
                    }
                }
            }
        }
    }
}

#[test]
fn guide_operators_compute_as_specified() {
    let env = Env::new(100.0, 50.0);
    let f = |fmla: &str| (env.formula(fmla) * 1e6).round() / 1e6;
    assert_eq!(f("*/ w 3 4"), 75.0);
    assert_eq!(f("+- w h 10"), 140.0);
    assert_eq!(f("+/ w h 3"), 50.0);
    assert_eq!(f("?: -1 7 9"), 9.0);
    assert_eq!(f("abs -5"), 5.0);
    assert_eq!(f("at2 1 1"), 2_700_000.0); // 45°
    assert_eq!(f("cos 10 cd4"), 0.0);
    assert_eq!(f("sin 10 cd4"), 10.0);
    assert_eq!(
        f("cat2 10 1 1"),
        (10.0 * (PI / 4.0).cos() * 1e6).round() / 1e6
    );
    assert_eq!(f("max 3 h"), 50.0);
    assert_eq!(f("min 3 h"), 3.0);
    assert_eq!(f("mod 3 4 0"), 5.0);
    assert_eq!(f("pin 0 120 w"), 100.0);
    assert_eq!(f("sqrt 16"), 4.0);
    assert_eq!(f("val ss"), 50.0);
    assert_eq!(f("*/ 1 2 0"), 0.0); // division by zero reads as 0
}
