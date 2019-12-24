package exchange

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"net/url"
	"strings"
	"time"

	"github.com/byuoitav/common/log"
	"github.com/byuoitav/scheduler/calendars"
)

type Calendar struct {
	ClientId     string
	ClientSecret string
	TennantId    string
	RoomID       string
	RoomResource string
}

// GetEvents will get all the days events in exchange
func (c *Calendar) GetEvents(ctx context.Context) ([]calendars.Event, error) {
	token, err := c.GetToken(ctx)
	if err != nil {
		return nil, fmt.Errorf("error getting exchange token: %w", err)
	}
	bearerToken := "Bearer " + token

	calendarID, err := c.GetCalendarID(ctx, token)
	if err != nil {
		return nil, fmt.Errorf("error getting calendar ID for room: %s, %w", c.RoomID, err)
	}

	reqURL := "https://outlook.office.com/api/v2.0/users/" + c.RoomResource + "/calendars/" + calendarID + "/calendarView"
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, reqURL, nil)
	if err != nil {
		return nil, fmt.Errorf("error creating get request to: %s, %w", reqURL, err)
	}
	req.Header.Add("Authorization", bearerToken)

	//Add query parameters
	loc, _ := time.LoadLocation("UTC")
	currentTime := time.Now()
	currentDayBeginning := time.Date(currentTime.Year(), currentTime.Month(), currentTime.Day(), 0, 0, 0, 0, currentTime.Location())
	currentDayEnding := time.Date(currentTime.Year(), currentTime.Month(), currentTime.Day(), 23, 59, 59, 0, currentTime.Location())
	currentDayBeginning = currentDayBeginning.In(loc)
	currentDayEnding = currentDayEnding.In(loc)

	query := req.URL.Query()
	query.Add("startDateTime", currentDayBeginning.Format("2006-01-02T15:04:05"))
	query.Add("endDateTime", currentDayEnding.Format("2006-01-02T15:04:05"))
	req.URL.RawQuery = query.Encode()

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return nil, fmt.Errorf("error sending http request to: %s, %w", reqURL, err)
	}
	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		log.L.Errorf("Error reading response body | %v", err)
		return nil, fmt.Errorf("Error reading response body | %v", err)
	}

	if resp.StatusCode/100 != 2 {
		return nil, fmt.Errorf("invalid response code %v: %s", resp.StatusCode, body)
	}

	var respBody eventResponse
	if err := json.Unmarshal(body, &respBody); err != nil {
		return nil, fmt.Errorf("error unmarshalling response body from json into exchange event object: %w", err)
	}

	dateTimeLayout := "2006-01-02T15:04:05"
	var events []calendars.Event
	for _, event := range respBody.Events {
		eventStart, err := time.Parse(dateTimeLayout, event.Start.DateTime)
		if err != nil {
			return nil, fmt.Errorf("error parsing exchange event start time into go time struct: %w", err)
		}
		eventEnd, err := time.Parse(dateTimeLayout, event.End.DateTime)
		if err != nil {
			return nil, fmt.Errorf("error parsing exchange event end time into go time struct: %w", err)
		}

		timeZone, _ := time.Now().Zone()
		location, _ := time.LoadLocation(timeZone)
		eventStart = eventStart.In(location)
		eventEnd = eventEnd.In(location)
		events = append(events, calendars.Event{
			Title:     event.Subject,
			StartTime: eventStart,
			EndTime:   eventEnd,
		})
	}

	return events, nil
}

// CreateEvent will create an exchange event
func (c *Calendar) CreateEvent(ctx context.Context, event calendars.Event) error {
	token, err := c.GetToken(ctx)
	if err != nil {
		return fmt.Errorf("error getting auth token: %w", err)
	}
	bearerToken := "Bearer " + token

	calendarID, err := c.GetCalendarID(ctx, token)
	if err != nil {
		return fmt.Errorf("error getting calendar ID for room: %s, %w", c.RoomID, err)
	}

	//Convert calendar event into exchange event
	loc, _ := time.LoadLocation("UTC")

	eventStart := event.StartTime.In(loc)

	eventEnd := event.EndTime.In(loc)

	reqBodyStruct := eventRequest{
		Subject: event.Title,
		Body: eventBody{
			ContentType: "HTML",
			Content:     "",
		},
		Start: exchangeDate{
			DateTime: eventStart.Format("2006-01-02T15:04:05"),
			TimeZone: "Etc/GMT",
		},
		End: exchangeDate{
			DateTime: eventEnd.Format("2006-01-02T15:04:05"),
			TimeZone: "Etc/GMT",
		},
		Attendees: make([]attendee, 0),
	}
	reqBodyString, err := json.Marshal(reqBodyStruct)
	if err != nil {
		return fmt.Errorf("error converting request body to json string: %w", err)
	}

	reqURL := "https://outlook.office.com/api/v2.0/users/" + c.RoomResource + "/calendars/" + calendarID + "/events"
	req, err := http.NewRequestWithContext(ctx, http.MethodPost, reqURL, bytes.NewBuffer(reqBodyString))
	if err != nil {
		return fmt.Errorf("error creating post request to: %s, %w", reqURL, err)
	}

	req.Header.Add("Authorization", bearerToken)
	req.Header.Add("Content-Type", "application/json")

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return fmt.Errorf("error sending http request to: %s, %w", reqURL, err)
	}
	defer resp.Body.Close()

	if resp.StatusCode/100 != 2 {
		body, err := ioutil.ReadAll(resp.Body)
		if err != nil {
			return err
		}

		return fmt.Errorf("invalid response code %v: %s", resp.StatusCode, body)
	}

	return nil
}

// GetCalendarID will get the exchange calendar id
func (c *Calendar) GetCalendarID(ctx context.Context, token string) (string, error) {
	reqURL := "https://outlook.office.com/api/v2.0/users/" + c.RoomResource + "/calendars"
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, reqURL, nil)
	if err != nil {
		return "", fmt.Errorf("error creating get request to: %s, %w", reqURL, err)
	}

	req.Header.Add("Authorization", "Bearer "+token)

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return "", fmt.Errorf("error sending http request to: %s, %w", reqURL, err)
	}
	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return "", fmt.Errorf("error reading response body: %w", err)
	}

	if resp.StatusCode/100 != 2 {
		return "", fmt.Errorf("invalid response code %v: %s", resp.StatusCode, body)
	}

	var respBody calendarResponse
	if err := json.Unmarshal(body, &respBody); err != nil {
		return "", fmt.Errorf("error unmarshalling response body from json into exchange calendar object: %w", err)
	}

	// Locate the proper calendar
	if len(respBody.Calendars) == 0 {
		return "", fmt.Errorf("there are no calendars listed for this resource")
	}
	if len(respBody.Calendars) > 1 {
		for _, calendar := range respBody.Calendars {
			if c.RoomID == calendar.Name {
				return calendar.ID, nil
			}
		}
	}
	//Return id of first calendar by default
	return respBody.Calendars[0].ID, nil
}

// GetToken will get an exchange oauth token
func (c *Calendar) GetToken(ctx context.Context) (string, error) {
	bodyParams := url.Values{}
	bodyParams.Set("client_id", c.ClientId)
	bodyParams.Set("scope", "https://outlook.office.com/.default")
	bodyParams.Set("client_secret", c.ClientSecret)
	bodyParams.Set("grant_type", "client_credentials")

	reqURL := "https://login.microsoftonline.com/" + c.TennantId + "/oauth2/v2.0/token"

	req, err := http.NewRequestWithContext(ctx, http.MethodPost, reqURL, strings.NewReader(bodyParams.Encode()))
	req.Header.Set("Content-type", "application/x-www-form-urlencoded")
	if err != nil {
		return "", fmt.Errorf("cannot make HTTP Post request to: %s, %w", reqURL, err)
	}

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return "", fmt.Errorf("cannot send request to: %s, %w", reqURL, err)
	}
	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return "", fmt.Errorf("error resolving response body: %w", err)
	}

	if resp.StatusCode/100 != 2 {
		return "", fmt.Errorf("invalid response code %v: %s", resp.StatusCode, body)
	}

	var respBody token
	if err := json.Unmarshal([]byte(body), &respBody); err != nil {
		return "", fmt.Errorf("error unmarshalling json body: %w", err)
	}

	return respBody.Token, nil
}
