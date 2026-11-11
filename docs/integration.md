# Student Selector Integration

The selector is designed to run either as its own GitHub Pages app or as an on-demand widget inside another lesson page.

Published base URL:

```text
https://randomizerselection.github.io/studentselector/
```

## Standalone Link

Use this when a lesson page should open the selector in a separate tab.

```html
<a href="https://randomizerselection.github.io/studentselector/" target="_blank" rel="noopener">
  Open student selector
</a>
```

## Built-In Overlay

Use this when the host page can accept the selector's own overlay and stylesheet.

```html
<script src="https://randomizerselection.github.io/studentselector/selector.js"></script>
<button type="button" id="open-selector">Student selector</button>

<script>
  document.getElementById("open-selector").addEventListener("click", () => {
    window.StudentSelector.open({
      basePath: "https://randomizerselection.github.io/studentselector/"
    });
  });
</script>
```

`open()` returns:

```js
{
  app,      // StudentSelectorApp instance
  close,    // closes and removes the overlay
  element   // overlay host element
}
```

## Mount In A Host Panel

Use this for lesson-slide integration where the selector should appear in a right-side panel and should not cover the whole slide.

```html
<aside id="student-selector-panel"></aside>
<script src="https://randomizerselection.github.io/studentselector/selector.js"></script>
<script>
  const panel = document.getElementById("student-selector-panel");

  const selector = window.StudentSelector.mount(panel, {
    basePath: "https://randomizerselection.github.io/studentselector/"
  });
</script>
```

`mount()` returns the app instance. Call `destroy()` when the host removes the panel:

```js
selector.destroy();
```

## Host-Supplied Styles

If the host page needs scoped styles, set `skipStyles: true` and provide your own CSS for the selector classes inside the host container.

```js
window.StudentSelector.mount(panel, {
  basePath: "https://randomizerselection.github.io/studentselector/",
  skipStyles: true,
  onClose: () => panel.remove()
});
```

This is the recommended pattern for the IGCSE lesson slide viewer because it prevents the standalone selector stylesheet from affecting slide typography or colors.

## Assets And State

The selector loads these files relative to `basePath`:

- `assets/messages.csv`
- `assets/icon.png`
- optional audio files in `assets/`

Choose **Load roster CSV** inside the selector to read a local roster. No roster is fetched from the site or uploaded by the selector. Roster and selection state are stored in `sessionStorage`, scoped to the browser tab and the host page's origin. Only embed the selector in trusted pages, since scripts on the host page can access its data. Close the tab when finished on a shared computer. Browser session restoration may restore tab data.

