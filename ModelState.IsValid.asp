public async Task<IActionResult> Create(Movie movie) // Noncompliant: model validity check is missing
{
    _context.Movies.Add(movie);
    await _context.SaveChangesAsync();

    return RedirectToAction(nameof(Index));
}