// Contribution Tracker 2026 – small helpers

document.addEventListener('DOMContentLoaded', function () {
    // Auto-dismiss flash messages after 5 seconds
    document.querySelectorAll('.flash').forEach(function (el) {
        setTimeout(function () {
            el.style.transition = 'opacity 0.4s';
            el.style.opacity = '0';
            setTimeout(function () { el.remove(); }, 400);
        }, 5000);
    });
});
