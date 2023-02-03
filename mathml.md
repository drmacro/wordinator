# Math support in Wordinator

The SWPX schema by default does not allow MathML to be included,
because the Wordinator does not support it out of the box. However,
the `src/main/doctypes/simplewpml/simplewpml.rng` schema is just a
shell that includes `simplewpml-base.rng`, which has no MathML
support. If you modify it to instead include `simplewpml-mathmlX.rng`
(where X is 2 or 3) you get a schema supporting inclusion of MathML
version X.

For MathML embedded in SWPX to actually turn into formulas in the Word
documents created by the Wordinator you need to supply an XSLT
stylesheet that converts MathML to the OOXML math format. (If you
don't have such a stylesheet, see the "Finding a stylesheet" section
below.)

## Running the Wordinator with the stylesheet

The stylesheet must be named `MML2OMML.XSL`, and will be loaded from
classpath.

If you run the Wordinator with `-jar wordinator.jar` *only* resources
in the jar file will be loaded, and so you will need to build a jar
file with the stylesheet included.

Alternatively, you can run the Wordinator as follows:

```
java -cp wordinator.jar:path/to/directory/with/stylesheet org.wordinator.xml2docx.MakeDocx ...
```

Running Java this way allows you to point to the directory where the
stylesheet is, enabling Java to find it.

## Finding a stylesheet

No open source MathML to OOXML stylesheet appears to exist, but one is
distributed with Microsoft Office. For copyright reasons it cannot be
installed in the Wordinator, but you can extract it from a Microsoft
Office installation and make it available to the Wordinator.

You can find this stylesheet in a Windows installation of Microsoft
Word at the following location:

```
C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL
```

You can of course supply any other stylesheet as you wish.

## Using MathML in SWPX

Once you've turned on MathML support there are two places where
MathML's `<mml:math xmlns:mml="http://www.w3.org/1998/Math/MathML">`
element can be used in an SWPX file:

  * as a child of `<wp:p>` for block-level formulae, and
  * as a child of `<wp:run>` for inline-level formulae.

See the test file `test/simplewpml-test-mathml-01.xml` for an example
of both placements.

## Schemas

By default, the `simplewpml.rng` schema points to a combination of the base SWPX schema and a modified MathML3 schema. It can be changed readily to support a modified MathML2 schema. These MathML RNG schemas were created from the DTDs and then modified specifically for use in Wordinator validation where annotations are not permitted.

Typical MathML annotations are permitted to be restricted and/or augmented by users. These are important constructs in user-facing MathML markup. The MML2OMML.XSL stylesheet does not accommodate annotations and the MathML markup used with SWPX never is user-facing, and so can be deleted from the MathML stream without consequence. The absence of handling annotations in the MML2OMML.XSL stylesheet requires annotations to be deleted.
