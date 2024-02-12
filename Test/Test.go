package main

import (
	"context"
    "log"
	"time"
	"fmt"

    "github.com/chromedp/chromedp"
	//"github.com/chromedp/cdproto/dom"
	"github.com/xuri/excelize/v2"
)

func LogOut(ctx context.Context) {
	if err := chromedp.Run(ctx, 
		// -------------------------- Log Out ------------------------------
		chromedp.Click(".user_menu_icon"),
		chromedp.Click(".menu_link.icon-logout"),
	) 
		
	err != nil {
		log.Fatal(err)
	}
}

func main() { // STILL WORKS
	opts := append(chromedp.DefaultExecAllocatorOptions[:],
		chromedp.Flag("headless", false),
		chromedp.Flag("disable-gpu", false),
		chromedp.Flag("enable-automation", false),
		chromedp.Flag("disable-extensions", false),
	)
	
	allocCtx, cancel := chromedp.NewExecAllocator(context.Background(), opts...)
	defer cancel()
	
	// create context
	ctx, cancel := chromedp.NewContext(allocCtx, chromedp.WithLogf(log.Printf))
	defer cancel()

	// Open up file
	file, err := excelize.OpenFile("C:/Users/carla/Downloads/179972 - Values - NCOA updates.xlsx") // https://www.kelche.co/blog/go/excel/
	if err != nil {
			log.Fatal(err)
	}

	// Get all row values
	rows, err := file.GetRows("179972  Values  NCOA updates")
	if err != nil {
			log.Fatal(err)
	}
	fmt.Printf("Row: %d\n", len(rows))

	// Get all column values
	columns, err := file.GetCols("179972  Values  NCOA updates")
	if err != nil {
			log.Fatal(err)
	}
	fmt.Printf("Columns: %d\n", len(columns))

	// Get all Ref numbers
	// for _, row := range rows {
	// 	if len(row) > 0 { // Check if the row is not empty
	// 		fmt.Println(row[0]) // Print the value from the first column
	// 	}
	// }
	
	// Run chromedp
	if err := chromedp.Run(ctx,
		// -------------------------- Enter School Data for Website ------------------------------
		chromedp.Navigate("https://discover.highpoint.edu/manage/"),
		chromedp.WaitVisible("#userNameInput"),
		chromedp.Sleep(1000*time.Millisecond), // Wait 1 second
		chromedp.SendKeys("#userNameInput", "clapiz"),
		chromedp.SendKeys("#passwordInput", "Mich7890!@#$"),
		chromedp.Sleep(1000*time.Millisecond),
		chromedp.Click("#submitButton", chromedp.ByQuery), // Use cssSelector value (use SelectorsHub)
		chromedp.Sleep(2000*time.Millisecond),
	)

	err != nil {
		log.Fatal(err)
	}

	// for i := 0; i < len(rows); i++ {
	// 	fmt.Print("")
	// }

	// Put some sort of UI here

	LogOut(ctx)
	
	// prevent it from closing the browser immediately
	time.Sleep(5000*time.Millisecond)

	fmt.Println("Done")
}