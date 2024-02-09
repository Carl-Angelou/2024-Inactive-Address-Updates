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

func main() {
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

	file, err := excelize.OpenFile("students.xlsx")
	if err != nil {
			log.Fatal(err)
	}
	
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

		// -------------------------- Enter School Data for Website ------------------------------
		chromedp.WaitVisible("#footer_school"),


	)
	
	err != nil {
		log.Fatal(err)
	}

	fmt.Println("Done")
}