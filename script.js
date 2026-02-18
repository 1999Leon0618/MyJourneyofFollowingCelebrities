// Sticky Navbar Effect
const navbar = document.getElementById('navbar');

// USER PROVIDED URL
const GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ1iHfm2TwBzjzR9d72k98CfyKF8lDlAfFDCuv8A_SLeMfIofhtczuliDnFVemB22vnt_lPOhBSU3eD/pubhtml';

window.addEventListener('scroll', () => {
    if (window.scrollY > 50) {
        navbar.classList.add('scrolled');
    } else {
        navbar.classList.remove('scrolled');
    }
});

// Smooth Scroll for Anchor Links
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();

        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            target.scrollIntoView({
                behavior: 'smooth'
            });
        }
    });
});

// Intersection Observer for Fade-in Animations
const observerOptions = {
    threshold: 0.1
};

const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
        if (entry.isIntersecting) {
            entry.target.style.opacity = '1';
            entry.target.style.transform = 'translateY(0)';
            observer.unobserve(entry.target);
        }
    });
}, observerOptions);

// Apply animations to sections
document.querySelectorAll('.profile-card, .gallery-item, .timeline-item').forEach(el => {
    el.style.opacity = '0';
    el.style.transform = 'translateY(30px)';
    el.style.transition = 'opacity 0.6s ease-out, transform 0.6s ease-out';
    observer.observe(el);
});

// --- Google Sheets XLSX Integration ---

function getExportUrl(url) {
    // Handle "Published to Web" URL
    // From: .../pubhtml
    // To:   .../pub?output=xlsx
    if (url.includes('/pubhtml')) {
        return url.replace('/pubhtml', '/pub?output=xlsx');
    }

    // Handle "Edit" URL
    // From: .../edit...
    // To:   .../export?format=xlsx
    if (url.includes('/edit')) {
        return url.replace(/\/edit.*/, '/export?format=xlsx');
    }

    return url;
}

async function fetchAndParseXlsx() {
    const timelineContainer = document.getElementById('timeline-container');
    if (!timelineContainer) return;

    try {
        const xlsxUrl = getExportUrl(GOOGLE_SHEET_URL);
        console.log('Fetching XLSX from:', xlsxUrl);

        timelineContainer.innerHTML = '<div class="timeline-item"><div class="timeline-content"><p>Loading records from Google Sheets...</p></div></div>';

        const response = await fetch(xlsxUrl);
        if (!response.ok) throw new Error('Network response was not ok: ' + response.statusText);

        const buffer = await response.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.worksheets[0]; // Assume data is in first sheet
        const events = [];

        // Iterate rows (skip header, so start from row 2)
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header

            // Cell columns: 1:Date, 2:Artist, 3:EventName, 4:Location, 5:Seat, 6:Image URL
            const getCellValue = (idx) => {
                const cell = row.getCell(idx);
                const val = cell.value;
                if (val && typeof val === 'object') {
                    if (val.text) return val.text;
                    if (val.result) return val.result;
                    if (val.hyperlink) return val.text || val.hyperlink;
                }
                return val ? val.toString() : '';
            };

            // Date handling
            let dateVal = row.getCell(1).value;
            let dateStr = '';
            if (dateVal instanceof Date) {
                // Use UTC methods to avoid timezone shifts from Excel parsing
                const y = dateVal.getUTCFullYear();
                const m = String(dateVal.getUTCMonth() + 1).padStart(2, '0');
                const d = String(dateVal.getUTCDate()).padStart(2, '0');

                // Add time if present (non-zero) or if explicitly needed
                const hh = dateVal.getUTCHours();
                const mm = dateVal.getUTCMinutes();

                // If it looks like a full timestamp (not 00:00), show time
                // Adjust logic: sometimes excel dates are just dates (00:00)
                if (hh !== 0 || mm !== 0) {
                    dateStr = `${y}.${m}.${d} ${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
                } else {
                    dateStr = `${y}.${m}.${d}`;
                }
            } else if (typeof dateVal === 'object' && dateVal.result) {
                // Formula result
                dateStr = String(dateVal.result);
            } else {
                dateStr = dateVal ? String(dateVal) : '';
            }

            // Basic validation
            if (!dateStr || dateStr.trim() === '') return;

            // Helper to convert Google Drive links to direct content URLs
            // Using lh3.googleusercontent.com/d/{id} often works better for direct images
            const convertDriveLink = (url) => {
                if (!url) return url;
                // Handle standard View/Preview links: https://drive.google.com/file/d/ID/view...
                const fileMatch = url.match(/\/file\/d\/([-\w]{25,})/);
                if (fileMatch && fileMatch[1]) {
                    // return `https://drive.google.com/uc?export=view&id=${fileMatch[1]}`;
                    return `https://lh3.googleusercontent.com/d/${fileMatch[1]}`;
                }

                // Handle other drive links
                if (url.includes('drive.google.com') || url.includes('docs.google.com')) {
                    // Extract ID using a broader regex if not caught above
                    const idMatch = url.match(/[-\w]{25,}/);
                    if (idMatch) {
                        return `https://lh3.googleusercontent.com/d/${idMatch[0]}`;
                    }
                }
                return url;
            };

            // Get Image URL from column 6
            let rawImg = getCellValue(6).trim();
            // console.log(`Row ${rowNumber}: Raw Image Cell Value: "${rawImg}"`);

            let imgUrl = rawImg;

            // Smart Path Handling
            if (imgUrl.length > 0) {
                if (imgUrl.match(/^https?:\/\//i)) {
                    // It's a web URL, check if drive link
                    imgUrl = convertDriveLink(imgUrl);
                    // console.log(`Row ${rowNumber}: Converted URL: "${imgUrl}"`);
                } else if (!imgUrl.includes('/')) {
                    // Assume local file if no slashes
                    imgUrl = `images/${imgUrl}`;
                    // console.log(`Row ${rowNumber}: Local Path: "${imgUrl}"`);
                }
            }

            events.push({
                date: dateStr,
                artist: getCellValue(2),
                eventName: getCellValue(3),
                location: getCellValue(4),
                seat: getCellValue(5),
                imageUrl: imgUrl
            });
        });

        // Sort by Date Descending (Newest First)
        // This implements "倒敘"
        events.sort((a, b) => {
            const da = new Date(a.date.replace(/\./g, '-').replace(/\//g, '-'));
            const db = new Date(b.date.replace(/\./g, '-').replace(/\//g, '-'));
            if (isNaN(da)) return 1;
            if (isNaN(db)) return -1;
            return db - da; // Descending: B (New) - A (Old) > 0 -> B comes first
        });

        // Render
        timelineContainer.innerHTML = '';

        if (events.length === 0) {
            timelineContainer.innerHTML = '<div class="timeline-item"><div class="timeline-content"><p>No events found in the sheet.</p></div></div>';
            return;
        }

        events.forEach(event => {
            const item = document.createElement('div');
            item.className = 'timeline-item';

            let htmlContent = `
                <div class="timeline-dot"></div>
                <div class="timeline-content">
                    <span class="date">${event.date}</span>
                    <h4>${event.artist}</h4>
                    <p>${event.eventName}</p>
            `;

            if (event.location && event.location !== 'undefined') htmlContent += `<p>${event.location}</p>`;

            // Seat Visibility Logic
            try {
                const dateStr = event.date.replace(/\./g, '-').replace(/\//g, '-');
                const revealDate = new Date(dateStr);
                revealDate.setDate(revealDate.getDate() + 7);
                const now = new Date();

                if (event.seat && event.seat !== 'undefined') {
                    if (now >= revealDate) {
                        htmlContent += `<p>${event.seat}</p>`;
                    }
                }
            } catch (e) { console.warn(e); }

            // Image
            if (event.imageUrl && event.imageUrl !== 'images/') {
                // console.log(`Loading image for ${event.eventName}: ${event.imageUrl}`); // Debug log
                // Add error handling for image
                const img = document.createElement('img');
                img.src = event.imageUrl;
                img.style.width = '100%';
                img.style.borderRadius = '10px';
                img.style.marginTop = '10px';
                img.onerror = () => {
                    console.error(`Failed to load image: ${event.imageUrl}`);
                    img.style.display = 'none'; // Hide if broken
                    const errorMsg = document.createElement('p');
                    errorMsg.style.color = 'red';
                    errorMsg.style.fontSize = '0.8em';
                    errorMsg.textContent = `(Image not found: ${event.imageUrl})`;
                    item.querySelector('.timeline-content').appendChild(errorMsg);
                };
                htmlContent += img.outerHTML;
                // htmlContent += `<img src="${event.imageUrl}" style="width: 100%; border-radius: 10px; margin-top: 10px;">`;
            }

            htmlContent += `</div>`;
            item.innerHTML = htmlContent;

            // Animation
            item.style.opacity = '0';
            item.style.transform = 'translateY(30px)';
            item.style.transition = 'opacity 0.6s ease-out, transform 0.6s ease-out';
            observer.observe(item);

            timelineContainer.appendChild(item);
        });

    } catch (error) {
        console.error('Error fetching/parsing XLSX:', error);
        timelineContainer.innerHTML = `<div class="timeline-item"><div class="timeline-content"><p style="color:red;">Error loading data: ${error.message}</p></div></div>`;
    }
}

// Auto-load on page start
fetchAndParseXlsx();
updateProfileImage();

function updateProfileImage() {
    const profileImg = document.querySelector('.bento-avatar');
    if (profileImg) {
        const originalSrc = profileImg.getAttribute('src');
        if (originalSrc && (originalSrc.includes('drive.google.com') || originalSrc.includes('docs.google.com'))) {
            // Re-use the conversion logic (simplified duplicate here or move helper out)
            // Since convertDriveLink is inside the other function, we'll duplicate the simple regex here 
            // or move convertDriveLink to global scope (better).
            // Let's use the lh3 format directly here for now to be safe.

            const fileMatch = originalSrc.match(/\/file\/d\/([-\w]{25,})/);
            if (fileMatch && fileMatch[1]) {
                const newSrc = `https://lh3.googleusercontent.com/d/${fileMatch[1]}`;
                profileImg.src = newSrc;
            }
        }
    }
}
