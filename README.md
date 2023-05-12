<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/sangeethnandakumar/Twileloop.UOW">
    <img src="https://iili.io/HeD5SzG.png" alt="Logo" width="80" height="80">
  </a>

  <h2 align="center"> Twileloop UOW (With LiteDB) </h2>
  <h4 align="center"> Free | Open-Source | Fast </h4>
</div>

## About
A lightweight and ready-made implementation of unit of work pattern + LiteDB. With ability to use multiple LiteDB databases, ready-made CRUD operations repository, LiteDB's native transactions and thread-safe features.

A plug & play package where you don't need to write lot and lot and lots of code to setp repositories and stuff. It's easy as 2 steps below

## License
> Twileloop.UOW is licensed under the MIT License. See the LICENSE file for more details.

#### This library is absolutely free. If it gives you a smile, A small coffee would be a great way to support my work. Thank you for considering it!
[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/sangeethnanda)

# Usage

## 1. Register all databases
```csharp
builder.Services.AddUnitOfWork((uow) => {
    uow.Connections = new List<LiteDBConnection>
    {
        new LiteDBConnection("<DB_NAME_1>", "Filename=DatabaseA.db; Mode=Shared; Password=****;"),
        new LiteDBConnection("<DB_NAME_2>", "Filename=DatabaseB.db; Mode=Shared; Password=****;")
    };
});
```

## 2. Inject and Use as required
```csharp
    [ApiController]
    public class HomeController : ControllerBase 
    {
        private readonly UnitOfWork uow;

        public HomeController(UnitOfWork uow)
        {
            this.uow = uow;
        }

        [HttpGet]
        public IActionResult Get() 
        {            
            try
            {
                // Step 1: Point to a database
                uow.UseDatabase("<DB_NAME>");

                //Step 2: Get a repository for your model 'Dogs'
                var dogRepo = uow.GetRepository<Dogs>();

                //Step 3: Do some fetch
                allDogs = dogRepo.GetAll().ToList();

                //Step 4: Or any CRUD operations you like
                dogRepo.Add(new Dog());

                //Step 5: Finally, commit or rollback if transaction need to maintain. That's it
                uow.Commit();

                return Ok(allDogs);
            }
            catch(Exception)
            {
                uow.Rollback();
            }            
        }

    }
```
