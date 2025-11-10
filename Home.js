document.addEventListener('DOMContentLoaded', () => {
    const hamburgerBtn = document.querySelector('.hamburger-menu');
    const closeMenuBtn = document.querySelector('.close-menu-btn');
    const mobileMenuOverlay = document.querySelector('.mobile-menu-overlay');
    const mobileDropdowns = document.querySelectorAll('.mobile-dropdown > a');

    // Toggle mobile menu visibility
    hamburgerBtn.addEventListener('click', () => {
        mobileMenuOverlay.classList.add('open');
        document.body.style.overflow = 'hidden'; // Prevent scrolling when menu is open
    });

    closeMenuBtn.addEventListener('click', () => {
        mobileMenuOverlay.classList.remove('open');
        document.body.style.overflow = ''; // Restore scrolling
    });

    // Toggle mobile submenu visibility
    mobileDropdowns.forEach(dropdown => {
        dropdown.addEventListener('click', (e) => {
            e.preventDefault(); // Prevent default link behavior
            const submenu = dropdown.nextElementSibling;
            if (submenu && submenu.classList.contains('mobile-submenu')) {
                submenu.classList.toggle('open');
                // Rotate arrow icon
                const icon = dropdown.querySelector('.fa-chevron-down');
                if (icon) {
                    icon.style.transform = submenu.classList.contains('open') ? 'rotate(180deg)' : 'rotate(0deg)';
                }
            }
        });
    });

    // Close mobile menu if a link is clicked (optional, good for user experience)
    const mobileNavLinks = document.querySelectorAll('.mobile-nav a:not(.mobile-dropdown > a)');
    mobileNavLinks.forEach(link => {
        link.addEventListener('click', () => {
            mobileMenuOverlay.classList.remove('open');
            document.body.style.overflow = '';
        });
    });




     ///impact stats

   function animateCounters() {
    const counters = document.querySelectorAll('.stat-number');
    counters.forEach(counter => {
      const target = +counter.getAttribute('data-target');
      const isCurrency = counter.textContent.startsWith('R');
      const hasPlus = counter.textContent.includes('+');
      let current = 0;
      const speed = 60;

      const update = () => {
        const increment = target / speed;
        current += increment;

        if (current < target) {
          counter.textContent = isCurrency
            ? 'R' + Math.floor(current).toLocaleString()
            : Math.floor(current).toLocaleString()+ (hasPlus ? '+' : '');
          requestAnimationFrame(update);
        } else {
          counter.textContent = isCurrency
            ? 'R' + target.toLocaleString()
            : target.toLocaleString()+ (hasPlus ? '+' : '');
        }
      };

      update();
    });
  }

  function isInViewport(element) {
    const rect = element.getBoundingClientRect();
    return rect.top <= window.innerHeight && rect.bottom >= 0;
  }

  let hasAnimated = false;
  window.addEventListener('scroll', () => {
    const section = document.querySelector('#impact-stats');
    if (!hasAnimated && isInViewport(section)) {
      animateCounters();
      hasAnimated = true;
    }
  });




//Staff Javascript
  const staffCards = document.querySelectorAll('.staff-card');

  staffCards.forEach(card => {
    card.addEventListener('click', () => {
      card.classList.toggle('active');
    });
  });

      
//Slideshow of quotes of donors and beneficiaries
  const section = document.querySelector('.slideshow-wrapper');
if (section) {   // <-- prevents script breaking when slideshow not found
  const track = section.querySelector('.slideshow-track');
  const cards = section.querySelectorAll('.student-card');
  const prev  = section.querySelector('.prev');
  const next  = section.querySelector('.next');

  let index = 0;

  function showSlide(i) {
    if (i >= cards.length) index = 0;
    else if (i < 0) index = cards.length - 1;
    else index = i;

    track.style.transform = `translateX(-${index * 100}%)`;
  }

  prev.addEventListener('click', () => showSlide(index - 1));
  next.addEventListener('click', () => showSlide(index + 1));

  showSlide(index);
  console.log({ track, cards, prev, next });
}

  //media and stories
 const tabButtons = document.querySelectorAll('.tab-btn');
            const impactHeader = document.querySelector('.impact-header');
            const initiativesHeader = document.querySelector('.initiatives-header');
            const mediaStoriesHeader = document.querySelector('.media-stories-header');
            
            tabButtons.forEach(button => {
                button.addEventListener('click', function() {
                    // Remove active class from all buttons and content
                    document.querySelectorAll('.tab-btn').forEach(btn => {
                        btn.classList.remove('active');
                    });
                    document.querySelectorAll('.tab-content').forEach(content => {
                        content.classList.remove('active');
                    });
                    
                    // Add active class to clicked button and corresponding content
                    this.classList.add('active');
                    const tabId = this.getAttribute('data-tab');
                    document.getElementById(tabId).classList.add('active');
                    
                    // Change header based on active tab
                    if (tabId === 'initiatives') {
                        impactHeader.classList.remove('media-stories');
                        impactHeader.classList.add('initiatives');
                        initiativesHeader.style.display = 'block';
                        mediaStoriesHeader.style.display = 'none';
                    } else if (tabId === 'media-stories') {
                        impactHeader.classList.remove('initiatives');
                        impactHeader.classList.add('media-stories');
                        initiativesHeader.style.display = 'none';
                        mediaStoriesHeader.style.display = 'block';
                    }
                });
            });
            
            // Initialize header
            initiativesHeader.style.display = 'block';
            mediaStoriesHeader.style.display = 'none';

            const hash = window.location.hash;

            if(hash == "#media-stories"){
              
                document.querySelectorAll(".tab-content").forEach(tab=> {
                  tab.classList.remove("active");
                });

                const tabtoact = document.querySelector(hash);

                if(tabtoact)
                {
                  tabtoact.classList.add("active");

                  document.querySelectorAll(".tab-btn").forEach(btn=>{
                    btn.classList.remove("active");
                    if(btn.dataset.tab === "media-stories"){
                      btn.classList.add("active");
                    }
                  });

                  const media = tabtoact.querySelector(".media-stories");
                  if(media){
                    setTimeout(()=> {
                      media.scrollIntoView({behavior:"instant",block:"start"});
                    },150);
                  }
                }
              }
 






});

async function loadDonors() {
  const response = await fetch("pictures/EIRDonations2.xlsx"); // path to your Excel file
  const arrayBuffer = await response.arrayBuffer();

  // Parse Excel
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]]; // first sheet
  const donors = XLSX.utils.sheet_to_json(sheet);

  const donorGrid = document.getElementById("donorGrid");

  donors.forEach(donor => {
    const name = donor["Paid from/Full name"];
    const project = donor["Project/Project name"];

    if (!name || !project) return; // skip blanks

    const card = document.createElement("div");
    card.className = "donor-card";
    card.innerHTML = `
      <img src="pictures/icon_page-0001.jpg" alt="icon" />
      <h3>${name}</h3>
      <p><strong>Project:</strong> ${project}</p>
    `;
    
    donorGrid.appendChild(card);
  });
}

loadDonors();

//days number

// function updateDaysLeft(){
//   const today=new Date();
//   const endofYear = new Date(today.getFullYear(),11,31);
//   const msPerDay = 1000*60*60*24;
//   const daysLeft = Math.ceil((endofYear-today)/ msPerDay);
//   document.getElementById("daysLeft").textContent = daysLeft;

// }

// updateDaysLeft();



//Campaings

 const slides = document.querySelectorAll('.hero-slide');
let currentSlide = 0;
let slideInterval = setInterval(showNextSlide, 3000); // auto-change

function showSlide(index) {
  slides[currentSlide].classList.remove('active');
  currentSlide = (index + slides.length) % slides.length; // wrap around
  slides[currentSlide].classList.add('active');
}

function showNextSlide() {
  showSlide(currentSlide + 1);
}

function showPrevSlide() {
  showSlide(currentSlide - 1);
}

// Arrows
document.querySelector('.arrow.right').addEventListener('click', () => {
  showNextSlide();
  resetInterval();
});

document.querySelector('.arrow.left').addEventListener('click', () => {
  showPrevSlide();
  resetInterval();
});

// Reset auto-slide after manual interaction
function resetInterval() {
  clearInterval(slideInterval);
  slideInterval = setInterval(showNextSlide, 3000);
}



//snow-effect
  //  const canvas = document.getElementById('snow-canvas');
  // const ctx = canvas.getContext('2d');
  // canvas.width = window.innerWidth;
  // canvas.height = window.innerHeight;

  // let flakes = [];

  // function createSnowflakes() {
  //   flakes = [];
  //   for (let i = 0; i < 100; i++) {
  //     flakes.push({
  //       x: Math.random() * canvas.width,
  //       y: Math.random() * canvas.height,
  //       radius: Math.random() * 4 + 1,
  //       density: Math.random() * 100
  //     });
  //   }
  // }

  // function drawSnowflakes() {
  //   ctx.clearRect(0, 0, canvas.width, canvas.height);
  //   ctx.fillStyle = "white";
  //   ctx.beginPath();
  //   for (let i = 0; i < flakes.length; i++) {
  //     const f = flakes[i];
  //     ctx.moveTo(f.x, f.y);
  //     ctx.arc(f.x, f.y, f.radius, 0, Math.PI * 2, true);
  //   }
  //   ctx.fill();
  //   moveSnowflakes();
  // }

  // let angle = 0;

  // function moveSnowflakes() {
  //   angle += 0.01;
  //   for (let i = 0; i < flakes.length; i++) {
  //     const f = flakes[i];
  //     f.y += Math.cos(angle + f.density) + 1 + f.radius / 2;
  //     f.x += Math.sin(angle) * 1;

  //     // Reset when off screen
  //     if (f.y > canvas.height) {
  //       flakes[i].y = 0;
  //       flakes[i].x = Math.random() * canvas.width;
  //     }
  //   }
  // }

  // function animateSnow() {
  //   drawSnowflakes();
  //   requestAnimationFrame(animateSnow);
  // }

  // createSnowflakes();
  // animateSnow();

  // // Update canvas on window resize
  // window.addEventListener('resize', () => {
  //   canvas.width = window.innerWidth;
  //   canvas.height = window.innerHeight;
  //   createSnowflakes();
  // });

  //Spring canvas

//  const canvasspring = document.getElementById("spring-canvas");


//   const stx = canvasspring.getContext("2d");

//   // Resize canvas to match its CSS size
//   function resizeCanvas() {
//     canvasspring.width = canvasspring.offsetWidth;
//     canvasspring.height = canvasspring.offsetHeight;
//   }
//   resizeCanvas();

//   const petals = [];
//   const colors = ["#FF1493", "#FFD700", "#FF69B4"]; // bright pink, yellow, hot pink

//   // Create petals
//   for (let i = 0; i < 60; i++) {
//     petals.push({
//       x: Math.random() * canvasspring.width,
//       y: Math.random() * canvasspring.height,
//       radiusX: 12 + Math.random() * 8, // larger for visibility
//       radiusY: 6 + Math.random() * 4,
//       speed: 1 + Math.random() * 3,
//       angle: Math.random() * 2 * Math.PI,
//       color: colors[Math.floor(Math.random() * colors.length)]
//     });
//   }

//   // Draw and animate petals
//   function drawPetals() {
//     stx.clearRect(0, 0, canvasspring.width, canvasspring.height);

//     petals.forEach(petal => {
//       stx.beginPath();
//       stx.ellipse(petal.x, petal.y, petal.radiusX, petal.radiusY, petal.angle, 0, 2 * Math.PI);
//       stx.fillStyle = petal.color;
//       stx.fill();

//       // Movement
//       petal.y += petal.speed;
//       petal.x += Math.sin(petal.angle) * 1;
//       petal.angle += 0.01;

//       // Reset if out of canvas
//       if (petal.y > canvasspring.height) {
//         petal.y = -10;
//         petal.x = Math.random() * canvasspring.width;
//       }
//     });

//     requestAnimationFrame(drawPetals);
//   }

//   drawPetals();

//   // Resize canvas on window resize
//   window.addEventListener("resize", () => {
//     resizeCanvas();
//   });

const sec = document.querySelector('.hero-section');
if (sec) {
  const rect = sec.getBoundingClientRect();
  console.log("Width:", rect.width);
  console.log("Height:", rect.height);
} else {
  console.log("Element not found");
}



  

